---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 83ebbad97df829b1ec64748ec6671ecf8f137496
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600302"
---
# <a name="item"></a><span data-ttu-id="e1e94-102">item</span><span class="sxs-lookup"><span data-stu-id="e1e94-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e1e94-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e1e94-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e1e94-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="e1e94-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-106">Requirements</span></span>

|<span data-ttu-id="e1e94-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-107">Requirement</span></span>|<span data-ttu-id="e1e94-108">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-110">1.0</span></span>|
|[<span data-ttu-id="e1e94-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="e1e94-112">Restricted</span></span>|
|[<span data-ttu-id="e1e94-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e1e94-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e1e94-115">Members and methods</span></span>

| <span data-ttu-id="e1e94-116">Membro	</span><span class="sxs-lookup"><span data-stu-id="e1e94-116">Member</span></span> | <span data-ttu-id="e1e94-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e1e94-118">attachments</span><span class="sxs-lookup"><span data-stu-id="e1e94-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="e1e94-119">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-119">Member</span></span> |
| [<span data-ttu-id="e1e94-120">bcc</span><span class="sxs-lookup"><span data-stu-id="e1e94-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="e1e94-121">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-121">Member</span></span> |
| [<span data-ttu-id="e1e94-122">body</span><span class="sxs-lookup"><span data-stu-id="e1e94-122">body</span></span>](#body-body) | <span data-ttu-id="e1e94-123">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-123">Member</span></span> |
| [<span data-ttu-id="e1e94-124">cc</span><span class="sxs-lookup"><span data-stu-id="e1e94-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="e1e94-125">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-125">Member</span></span> |
| [<span data-ttu-id="e1e94-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="e1e94-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e1e94-127">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-127">Member</span></span> |
| [<span data-ttu-id="e1e94-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e1e94-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e1e94-129">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-129">Member</span></span> |
| [<span data-ttu-id="e1e94-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e1e94-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e1e94-131">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-131">Member</span></span> |
| [<span data-ttu-id="e1e94-132">end</span><span class="sxs-lookup"><span data-stu-id="e1e94-132">end</span></span>](#end-datetime) | <span data-ttu-id="e1e94-133">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-133">Member</span></span> |
| [<span data-ttu-id="e1e94-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e1e94-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="e1e94-135">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-135">Member</span></span> |
| [<span data-ttu-id="e1e94-136">from</span><span class="sxs-lookup"><span data-stu-id="e1e94-136">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="e1e94-137">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-137">Member</span></span> |
| [<span data-ttu-id="e1e94-138">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="e1e94-138">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="e1e94-139">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-139">Member</span></span> |
| [<span data-ttu-id="e1e94-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e1e94-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e1e94-141">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-141">Member</span></span> |
| [<span data-ttu-id="e1e94-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="e1e94-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e1e94-143">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-143">Member</span></span> |
| [<span data-ttu-id="e1e94-144">itemId</span><span class="sxs-lookup"><span data-stu-id="e1e94-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e1e94-145">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-145">Member</span></span> |
| [<span data-ttu-id="e1e94-146">itemType</span><span class="sxs-lookup"><span data-stu-id="e1e94-146">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="e1e94-147">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-147">Member</span></span> |
| [<span data-ttu-id="e1e94-148">location</span><span class="sxs-lookup"><span data-stu-id="e1e94-148">location</span></span>](#location-stringlocation) | <span data-ttu-id="e1e94-149">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-149">Member</span></span> |
| [<span data-ttu-id="e1e94-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e1e94-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e1e94-151">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-151">Member</span></span> |
| [<span data-ttu-id="e1e94-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e1e94-152">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="e1e94-153">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-153">Member</span></span> |
| [<span data-ttu-id="e1e94-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e1e94-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="e1e94-155">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-155">Member</span></span> |
| [<span data-ttu-id="e1e94-156">organizer</span><span class="sxs-lookup"><span data-stu-id="e1e94-156">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="e1e94-157">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-157">Member</span></span> |
| [<span data-ttu-id="e1e94-158">recorrência</span><span class="sxs-lookup"><span data-stu-id="e1e94-158">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="e1e94-159">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-159">Member</span></span> |
| [<span data-ttu-id="e1e94-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e1e94-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="e1e94-161">Member</span><span class="sxs-lookup"><span data-stu-id="e1e94-161">Member</span></span> |
| [<span data-ttu-id="e1e94-162">sender</span><span class="sxs-lookup"><span data-stu-id="e1e94-162">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="e1e94-163">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-163">Member</span></span> |
| [<span data-ttu-id="e1e94-164">seriesid</span><span class="sxs-lookup"><span data-stu-id="e1e94-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="e1e94-165">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-165">Member</span></span> |
| [<span data-ttu-id="e1e94-166">start</span><span class="sxs-lookup"><span data-stu-id="e1e94-166">start</span></span>](#start-datetime) | <span data-ttu-id="e1e94-167">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-167">Member</span></span> |
| [<span data-ttu-id="e1e94-168">subject</span><span class="sxs-lookup"><span data-stu-id="e1e94-168">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="e1e94-169">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-169">Member</span></span> |
| [<span data-ttu-id="e1e94-170">to</span><span class="sxs-lookup"><span data-stu-id="e1e94-170">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="e1e94-171">Membro</span><span class="sxs-lookup"><span data-stu-id="e1e94-171">Member</span></span> |
| [<span data-ttu-id="e1e94-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e1e94-173">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-173">Method</span></span> |
| [<span data-ttu-id="e1e94-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e1e94-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="e1e94-175">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-175">Method</span></span> |
| [<span data-ttu-id="e1e94-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e1e94-177">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-177">Method</span></span> |
| [<span data-ttu-id="e1e94-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e1e94-179">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-179">Method</span></span> |
| [<span data-ttu-id="e1e94-180">close</span><span class="sxs-lookup"><span data-stu-id="e1e94-180">close</span></span>](#close) | <span data-ttu-id="e1e94-181">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-181">Method</span></span> |
| [<span data-ttu-id="e1e94-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e1e94-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="e1e94-183">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-183">Method</span></span> |
| [<span data-ttu-id="e1e94-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e1e94-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="e1e94-185">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-185">Method</span></span> |
| [<span data-ttu-id="e1e94-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="e1e94-187">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-187">Method</span></span> |
| [<span data-ttu-id="e1e94-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="e1e94-189">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-189">Method</span></span> |
| [<span data-ttu-id="e1e94-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="e1e94-190">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="e1e94-191">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-191">Method</span></span> |
| [<span data-ttu-id="e1e94-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e1e94-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="e1e94-193">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-193">Method</span></span> |
| [<span data-ttu-id="e1e94-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e1e94-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="e1e94-195">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-195">Method</span></span> |
| [<span data-ttu-id="e1e94-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="e1e94-197">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-197">Method</span></span> |
| [<span data-ttu-id="e1e94-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e1e94-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e1e94-199">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-199">Method</span></span> |
| [<span data-ttu-id="e1e94-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e1e94-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e1e94-201">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-201">Method</span></span> |
| [<span data-ttu-id="e1e94-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e1e94-203">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-203">Method</span></span> |
| [<span data-ttu-id="e1e94-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="e1e94-204">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="e1e94-205">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-205">Method</span></span> |
| [<span data-ttu-id="e1e94-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e1e94-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="e1e94-207">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-207">Method</span></span> |
| [<span data-ttu-id="e1e94-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="e1e94-209">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-209">Method</span></span> |
| [<span data-ttu-id="e1e94-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e1e94-211">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-211">Method</span></span> |
| [<span data-ttu-id="e1e94-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e1e94-213">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-213">Method</span></span> |
| [<span data-ttu-id="e1e94-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e1e94-215">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-215">Method</span></span> |
| [<span data-ttu-id="e1e94-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e1e94-217">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-217">Method</span></span> |
| [<span data-ttu-id="e1e94-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e1e94-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e1e94-219">Método</span><span class="sxs-lookup"><span data-stu-id="e1e94-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e1e94-220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-220">Example</span></span>

<span data-ttu-id="e1e94-221">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="e1e94-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e1e94-222">Membros</span><span class="sxs-lookup"><span data-stu-id="e1e94-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e1e94-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e1e94-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e1e94-224">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="e1e94-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e1e94-225">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-226">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="e1e94-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e1e94-227">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="e1e94-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-228">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-228">Type</span></span>

*   <span data-ttu-id="e1e94-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e1e94-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-230">Requirements</span></span>

|<span data-ttu-id="e1e94-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-231">Requirement</span></span>|<span data-ttu-id="e1e94-232">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-234">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-234">1.0</span></span>|
|[<span data-ttu-id="e1e94-235">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-236">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-238">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-239">Example</span></span>

<span data-ttu-id="e1e94-240">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e1e94-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e1e94-242">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e1e94-243">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="e1e94-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-244">Type</span></span>

*   [<span data-ttu-id="e1e94-245">Destinatários</span><span class="sxs-lookup"><span data-stu-id="e1e94-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e1e94-246">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-246">Requirements</span></span>

|<span data-ttu-id="e1e94-247">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-247">Requirement</span></span>|<span data-ttu-id="e1e94-248">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-249">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-250">1.1</span><span class="sxs-lookup"><span data-stu-id="e1e94-250">1.1</span></span>|
|[<span data-ttu-id="e1e94-251">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-252">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-253">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-254">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-255">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="e1e94-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="e1e94-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="e1e94-257">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-258">Type</span></span>

*   [<span data-ttu-id="e1e94-259">Body</span><span class="sxs-lookup"><span data-stu-id="e1e94-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="e1e94-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-260">Requirements</span></span>

|<span data-ttu-id="e1e94-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-261">Requirement</span></span>|<span data-ttu-id="e1e94-262">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-264">1.1</span><span class="sxs-lookup"><span data-stu-id="e1e94-264">1.1</span></span>|
|[<span data-ttu-id="e1e94-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-266">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-267">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-269">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-269">Example</span></span>

<span data-ttu-id="e1e94-270">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="e1e94-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="e1e94-271">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e1e94-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e1e94-273">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e1e94-274">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-275">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-275">Read mode</span></span>

<span data-ttu-id="e1e94-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-278">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-278">Compose mode</span></span>

<span data-ttu-id="e1e94-279">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-280">Type</span></span>

*   <span data-ttu-id="e1e94-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-282">Requirements</span></span>

|<span data-ttu-id="e1e94-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-283">Requirement</span></span>|<span data-ttu-id="e1e94-284">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-286">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-286">1.0</span></span>|
|[<span data-ttu-id="e1e94-287">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-288">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-289">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-290">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-290">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="e1e94-291">(anulável) conversationId :Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="e1e94-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="e1e94-292">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="e1e94-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e1e94-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e1e94-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-297">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-297">Type</span></span>

*   <span data-ttu-id="e1e94-298">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-299">Requirements</span></span>

|<span data-ttu-id="e1e94-300">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-300">Requirement</span></span>|<span data-ttu-id="e1e94-301">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-303">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-303">1.0</span></span>|
|[<span data-ttu-id="e1e94-304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-305">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-307">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="e1e94-309">dateTimeCreated :Data</span><span class="sxs-lookup"><span data-stu-id="e1e94-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="e1e94-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-312">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-312">Type</span></span>

*   <span data-ttu-id="e1e94-313">Data</span><span class="sxs-lookup"><span data-stu-id="e1e94-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-314">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-314">Requirements</span></span>

|<span data-ttu-id="e1e94-315">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-315">Requirement</span></span>|<span data-ttu-id="e1e94-316">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-317">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-318">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-318">1.0</span></span>|
|[<span data-ttu-id="e1e94-319">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-320">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-321">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-322">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-323">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="e1e94-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="e1e94-324">dateTimeModified :Date</span></span>

<span data-ttu-id="e1e94-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-327">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-328">Type</span></span>

*   <span data-ttu-id="e1e94-329">Data</span><span class="sxs-lookup"><span data-stu-id="e1e94-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-330">Requirements</span></span>

|<span data-ttu-id="e1e94-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-331">Requirement</span></span>|<span data-ttu-id="e1e94-332">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-334">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-334">1.0</span></span>|
|[<span data-ttu-id="e1e94-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-335">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-336">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-337">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-338">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e1e94-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e1e94-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e1e94-341">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="e1e94-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e1e94-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-344">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-344">Read mode</span></span>

<span data-ttu-id="e1e94-345">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-346">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-346">Compose mode</span></span>

<span data-ttu-id="e1e94-347">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e1e94-348">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="e1e94-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e1e94-349">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e1e94-350">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-350">Type</span></span>

*   <span data-ttu-id="e1e94-351">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e1e94-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-352">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-352">Requirements</span></span>

|<span data-ttu-id="e1e94-353">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-353">Requirement</span></span>|<span data-ttu-id="e1e94-354">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-355">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-356">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-356">1.0</span></span>|
|[<span data-ttu-id="e1e94-357">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-358">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-359">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-360">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-360">Compose or Read</span></span>|

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="e1e94-361">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="e1e94-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="e1e94-362">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-363">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-363">Read mode</span></span>

<span data-ttu-id="e1e94-364">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-365">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-365">Compose mode</span></span>

<span data-ttu-id="e1e94-366">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-367">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-367">Type</span></span>

*   [<span data-ttu-id="e1e94-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e1e94-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="e1e94-369">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-369">Requirements</span></span>

|<span data-ttu-id="e1e94-370">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-370">Requirement</span></span>|<span data-ttu-id="e1e94-371">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-372">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-373">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-373">Preview</span></span>|
|[<span data-ttu-id="e1e94-374">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-374">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-375">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-376">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-377">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-378">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-378">Example</span></span>

<span data-ttu-id="e1e94-379">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-379">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="e1e94-380">de:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e1e94-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="e1e94-381">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="e1e94-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-384">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-385">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-385">Read mode</span></span>

<span data-ttu-id="e1e94-386">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-387">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-387">Compose mode</span></span>

<span data-ttu-id="e1e94-388">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="e1e94-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-389">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-389">Type</span></span>

*   <span data-ttu-id="e1e94-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e1e94-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-391">Requirements</span></span>

|<span data-ttu-id="e1e94-392">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e1e94-393">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-394">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-394">1.0</span></span>|<span data-ttu-id="e1e94-395">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-395">1.7</span></span>|
|[<span data-ttu-id="e1e94-396">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-397">ReadItem</span></span>|<span data-ttu-id="e1e94-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-399">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-400">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-400">Read</span></span>|<span data-ttu-id="e1e94-401">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-401">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="e1e94-402">Internetheaders::[internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="e1e94-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="e1e94-403">Obtém ou define os cabeçalhos de Internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-404">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-404">Type</span></span>

*   [<span data-ttu-id="e1e94-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e1e94-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="e1e94-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-406">Requirements</span></span>

|<span data-ttu-id="e1e94-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-407">Requirement</span></span>|<span data-ttu-id="e1e94-408">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-410">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-410">Preview</span></span>|
|[<span data-ttu-id="e1e94-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-412">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-413">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-414">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="e1e94-416">internetMessageId Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="e1e94-416">internetMessageId :String</span></span>

<span data-ttu-id="e1e94-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-419">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-419">Type</span></span>

*   <span data-ttu-id="e1e94-420">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-421">Requirements</span></span>

|<span data-ttu-id="e1e94-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-422">Requirement</span></span>|<span data-ttu-id="e1e94-423">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-425">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-425">1.0</span></span>|
|[<span data-ttu-id="e1e94-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-427">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-429">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

#### <a name="itemclass-string"></a><span data-ttu-id="e1e94-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="e1e94-431">itemClass :String</span></span>

<span data-ttu-id="e1e94-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e1e94-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="e1e94-436">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-436">Type</span></span>|<span data-ttu-id="e1e94-437">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-437">Description</span></span>|<span data-ttu-id="e1e94-438">classe de item</span><span class="sxs-lookup"><span data-stu-id="e1e94-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="e1e94-439">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="e1e94-439">Appointment items</span></span>|<span data-ttu-id="e1e94-440">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="e1e94-441">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="e1e94-441">Message items</span></span>|<span data-ttu-id="e1e94-442">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="e1e94-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="e1e94-443">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-444">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-444">Type</span></span>

*   <span data-ttu-id="e1e94-445">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-446">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-446">Requirements</span></span>

|<span data-ttu-id="e1e94-447">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-447">Requirement</span></span>|<span data-ttu-id="e1e94-448">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-449">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-450">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-450">1.0</span></span>|
|[<span data-ttu-id="e1e94-451">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-451">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-452">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-453">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-453">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-454">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-455">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e1e94-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="e1e94-456">(nullable) itemId :String</span></span>

<span data-ttu-id="e1e94-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-459">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="e1e94-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e1e94-460">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e1e94-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e1e94-461">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e1e94-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e1e94-462">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="e1e94-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e1e94-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-465">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-465">Type</span></span>

*   <span data-ttu-id="e1e94-466">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-467">Requirements</span></span>

|<span data-ttu-id="e1e94-468">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-468">Requirement</span></span>|<span data-ttu-id="e1e94-469">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-471">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-471">1.0</span></span>|
|[<span data-ttu-id="e1e94-472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-473">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-474">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-475">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-476">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-476">Example</span></span>

<span data-ttu-id="e1e94-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="e1e94-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e1e94-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e1e94-480">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="e1e94-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e1e94-481">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-482">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-482">Type</span></span>

*   [<span data-ttu-id="e1e94-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e1e94-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e1e94-484">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-484">Requirements</span></span>

|<span data-ttu-id="e1e94-485">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-485">Requirement</span></span>|<span data-ttu-id="e1e94-486">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-487">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-488">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-488">1.0</span></span>|
|[<span data-ttu-id="e1e94-489">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-490">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-491">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-492">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-493">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="e1e94-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e1e94-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="e1e94-495">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-496">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-496">Read mode</span></span>

<span data-ttu-id="e1e94-497">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-498">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-498">Compose mode</span></span>

<span data-ttu-id="e1e94-499">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-500">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-500">Type</span></span>

*   <span data-ttu-id="e1e94-501">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e1e94-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-502">Requirements</span></span>

|<span data-ttu-id="e1e94-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-503">Requirement</span></span>|<span data-ttu-id="e1e94-504">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-506">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-506">1.0</span></span>|
|[<span data-ttu-id="e1e94-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-508">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-509">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-510">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-510">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e1e94-511">normalizedSubject :Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e1e94-511">normalizedSubject :String</span></span>

<span data-ttu-id="e1e94-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e1e94-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="e1e94-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-516">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-516">Type</span></span>

*   <span data-ttu-id="e1e94-517">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-518">Requirements</span></span>

|<span data-ttu-id="e1e94-519">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-519">Requirement</span></span>|<span data-ttu-id="e1e94-520">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-521">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-522">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-522">1.0</span></span>|
|[<span data-ttu-id="e1e94-523">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-524">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-525">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-526">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-527">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="e1e94-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="e1e94-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="e1e94-529">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-530">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-530">Type</span></span>

*   [<span data-ttu-id="e1e94-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e1e94-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="e1e94-532">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-532">Requirements</span></span>

|<span data-ttu-id="e1e94-533">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-533">Requirement</span></span>|<span data-ttu-id="e1e94-534">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-535">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-536">1.3</span><span class="sxs-lookup"><span data-stu-id="e1e94-536">1.3</span></span>|
|[<span data-ttu-id="e1e94-537">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-537">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-538">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-539">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-539">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-540">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-541">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-541">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e1e94-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e1e94-543">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="e1e94-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e1e94-544">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-545">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-545">Read mode</span></span>

<span data-ttu-id="e1e94-546">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-547">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-547">Compose mode</span></span>

<span data-ttu-id="e1e94-548">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-549">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-549">Type</span></span>

*   <span data-ttu-id="e1e94-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-551">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-551">Requirements</span></span>

|<span data-ttu-id="e1e94-552">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-552">Requirement</span></span>|<span data-ttu-id="e1e94-553">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-554">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-555">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-555">1.0</span></span>|
|[<span data-ttu-id="e1e94-556">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-556">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-557">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-558">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-558">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-559">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-559">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="e1e94-560">organizador:[](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e1e94-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="e1e94-561">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-562">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-562">Read mode</span></span>

<span data-ttu-id="e1e94-563">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-564">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-564">Compose mode</span></span>

<span data-ttu-id="e1e94-565">A `organizer` propriedade retorna um [](/javascript/api/outlook/office.organizer) objeto organizador que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="e1e94-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="e1e94-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-566">Type</span></span>

*   <span data-ttu-id="e1e94-567">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e1e94-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-568">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-568">Requirements</span></span>

|<span data-ttu-id="e1e94-569">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e1e94-570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-571">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-571">1.0</span></span>|<span data-ttu-id="e1e94-572">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-572">1.7</span></span>|
|[<span data-ttu-id="e1e94-573">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-574">ReadItem</span></span>|<span data-ttu-id="e1e94-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-576">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-576">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-577">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-577">Read</span></span>|<span data-ttu-id="e1e94-578">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-578">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="e1e94-579">(anulável) recorrência[](/javascript/api/outlook/office.recurrence) : recorrência</span><span class="sxs-lookup"><span data-stu-id="e1e94-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="e1e94-580">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="e1e94-581">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="e1e94-582">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="e1e94-583">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="e1e94-584">A `recurrence` propriedade retorna um [](/javascript/api/outlook/office.recurrence) objeto de recorrência para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="e1e94-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="e1e94-585">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="e1e94-586">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="e1e94-587">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="e1e94-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="e1e94-588">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="e1e94-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-589">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-589">Read mode</span></span>

<span data-ttu-id="e1e94-590">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="e1e94-591">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-592">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-592">Compose mode</span></span>

<span data-ttu-id="e1e94-593">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="e1e94-594">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-594">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e1e94-595">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-595">Type</span></span>

* [<span data-ttu-id="e1e94-596">Recorrência</span><span class="sxs-lookup"><span data-stu-id="e1e94-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="e1e94-597">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-597">Requirement</span></span>|<span data-ttu-id="e1e94-598">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-599">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-600">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-600">1.7</span></span>|
|[<span data-ttu-id="e1e94-601">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-601">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-602">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-603">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-603">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-604">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-604">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e1e94-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e1e94-606">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="e1e94-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e1e94-607">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-608">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-608">Read mode</span></span>

<span data-ttu-id="e1e94-609">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-610">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-610">Compose mode</span></span>

<span data-ttu-id="e1e94-611">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-612">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-612">Type</span></span>

*   <span data-ttu-id="e1e94-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-614">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-614">Requirements</span></span>

|<span data-ttu-id="e1e94-615">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-615">Requirement</span></span>|<span data-ttu-id="e1e94-616">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-617">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-618">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-618">1.0</span></span>|
|[<span data-ttu-id="e1e94-619">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-619">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-620">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-621">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-621">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-622">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-622">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="e1e94-623">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e1e94-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="e1e94-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e1e94-p129">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p129">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-628">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-629">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-629">Type</span></span>

*   [<span data-ttu-id="e1e94-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e1e94-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e1e94-631">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-631">Requirements</span></span>

|<span data-ttu-id="e1e94-632">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-632">Requirement</span></span>|<span data-ttu-id="e1e94-633">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-634">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-635">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-635">1.0</span></span>|
|[<span data-ttu-id="e1e94-636">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-636">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-637">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-638">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-638">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-639">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-640">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="e1e94-641">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="e1e94-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="e1e94-642">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="e1e94-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="e1e94-643">No OWA e no Outlook, `seriesId` o retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="e1e94-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="e1e94-644">No enTanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="e1e94-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-645">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="e1e94-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e1e94-646">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e1e94-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="e1e94-647">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e1e94-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e1e94-648">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="e1e94-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="e1e94-649">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="e1e94-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="e1e94-650">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-650">Type</span></span>

* <span data-ttu-id="e1e94-651">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-652">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-652">Requirements</span></span>

|<span data-ttu-id="e1e94-653">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-653">Requirement</span></span>|<span data-ttu-id="e1e94-654">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-655">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-656">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-656">1.7</span></span>|
|[<span data-ttu-id="e1e94-657">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-657">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-658">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-659">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-659">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-660">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-661">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-661">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e1e94-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e1e94-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e1e94-663">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="e1e94-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e1e94-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-666">Modo de Leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-666">Read mode</span></span>

<span data-ttu-id="e1e94-667">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-668">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-668">Compose mode</span></span>

<span data-ttu-id="e1e94-669">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e1e94-670">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="e1e94-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e1e94-671">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e1e94-672">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-672">Type</span></span>

*   <span data-ttu-id="e1e94-673">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e1e94-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-674">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-674">Requirements</span></span>

|<span data-ttu-id="e1e94-675">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-675">Requirement</span></span>|<span data-ttu-id="e1e94-676">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-677">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-678">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-678">1.0</span></span>|
|[<span data-ttu-id="e1e94-679">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-679">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-680">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-681">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-681">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-682">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-682">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="e1e94-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e1e94-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="e1e94-684">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e1e94-685">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="e1e94-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-686">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-686">Read mode</span></span>

<span data-ttu-id="e1e94-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="e1e94-689">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="e1e94-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-690">Modo de composição</span><span class="sxs-lookup"><span data-stu-id="e1e94-690">Compose mode</span></span>
<span data-ttu-id="e1e94-691">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-692">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-692">Type</span></span>

*   <span data-ttu-id="e1e94-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e1e94-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-694">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-694">Requirements</span></span>

|<span data-ttu-id="e1e94-695">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-695">Requirement</span></span>|<span data-ttu-id="e1e94-696">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-697">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-698">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-698">1.0</span></span>|
|[<span data-ttu-id="e1e94-699">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-699">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-700">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-701">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-701">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-702">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-702">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e1e94-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e1e94-704">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e1e94-705">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e1e94-706">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e1e94-706">Read mode</span></span>

<span data-ttu-id="e1e94-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="e1e94-709">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e1e94-709">Compose mode</span></span>

<span data-ttu-id="e1e94-710">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e1e94-711">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-711">Type</span></span>

*   <span data-ttu-id="e1e94-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e1e94-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-713">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-713">Requirements</span></span>

|<span data-ttu-id="e1e94-714">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-714">Requirement</span></span>|<span data-ttu-id="e1e94-715">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-716">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-717">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-717">1.0</span></span>|
|[<span data-ttu-id="e1e94-718">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-718">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-719">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-720">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-720">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-721">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e1e94-722">Métodos</span><span class="sxs-lookup"><span data-stu-id="e1e94-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e1e94-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e1e94-724">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e1e94-725">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="e1e94-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e1e94-726">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-727">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-727">Parameters</span></span>
|<span data-ttu-id="e1e94-728">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-728">Name</span></span>|<span data-ttu-id="e1e94-729">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-729">Type</span></span>|<span data-ttu-id="e1e94-730">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-730">Attributes</span></span>|<span data-ttu-id="e1e94-731">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="e1e94-732">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-732">String</span></span>||<span data-ttu-id="e1e94-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e1e94-735">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-735">String</span></span>||<span data-ttu-id="e1e94-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e1e94-738">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-738">Object</span></span>|<span data-ttu-id="e1e94-739">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-739">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-740">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-741">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-741">Object</span></span>|<span data-ttu-id="e1e94-742">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-742">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-743">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e1e94-744">Booliano</span><span class="sxs-lookup"><span data-stu-id="e1e94-744">Boolean</span></span>|<span data-ttu-id="e1e94-745">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-745">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-746">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e1e94-747">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-747">function</span></span>|<span data-ttu-id="e1e94-748">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-748">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-749">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e1e94-750">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e1e94-751">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="e1e94-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e1e94-752">Erros</span><span class="sxs-lookup"><span data-stu-id="e1e94-752">Errors</span></span>

|<span data-ttu-id="e1e94-753">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e1e94-753">Error code</span></span>|<span data-ttu-id="e1e94-754">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e1e94-755">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="e1e94-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e1e94-756">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="e1e94-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e1e94-757">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-758">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-758">Requirements</span></span>

|<span data-ttu-id="e1e94-759">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-759">Requirement</span></span>|<span data-ttu-id="e1e94-760">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-761">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-762">1.1</span><span class="sxs-lookup"><span data-stu-id="e1e94-762">1.1</span></span>|
|[<span data-ttu-id="e1e94-763">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-765">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-766">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e1e94-767">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e1e94-767">Examples</span></span>

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

<span data-ttu-id="e1e94-768">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="e1e94-769">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e1e94-770">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e1e94-771">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="e1e94-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="e1e94-772">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="e1e94-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="e1e94-773">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-774">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-774">Parameters</span></span>
|<span data-ttu-id="e1e94-775">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-775">Name</span></span>|<span data-ttu-id="e1e94-776">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-776">Type</span></span>|<span data-ttu-id="e1e94-777">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-777">Attributes</span></span>|<span data-ttu-id="e1e94-778">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="e1e94-779">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-779">String</span></span>||<span data-ttu-id="e1e94-780">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="e1e94-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="e1e94-781">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-781">String</span></span>||<span data-ttu-id="e1e94-p139">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e1e94-784">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-784">Object</span></span>|<span data-ttu-id="e1e94-785">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-785">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-786">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-787">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-787">Object</span></span>|<span data-ttu-id="e1e94-788">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-788">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-789">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e1e94-790">Booliano</span><span class="sxs-lookup"><span data-stu-id="e1e94-790">Boolean</span></span>|<span data-ttu-id="e1e94-791">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-791">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-792">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e1e94-793">função</span><span class="sxs-lookup"><span data-stu-id="e1e94-793">function</span></span>|<span data-ttu-id="e1e94-794">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-794">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-795">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e1e94-796">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e1e94-797">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="e1e94-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e1e94-798">Erros</span><span class="sxs-lookup"><span data-stu-id="e1e94-798">Errors</span></span>

|<span data-ttu-id="e1e94-799">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e1e94-799">Error code</span></span>|<span data-ttu-id="e1e94-800">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e1e94-801">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="e1e94-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e1e94-802">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="e1e94-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e1e94-803">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-804">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-804">Requirements</span></span>

|<span data-ttu-id="e1e94-805">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-805">Requirement</span></span>|<span data-ttu-id="e1e94-806">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-807">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-808">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-808">Preview</span></span>|
|[<span data-ttu-id="e1e94-809">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-811">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-812">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e1e94-813">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e1e94-813">Examples</span></span>

```javascript
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e1e94-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e1e94-815">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="e1e94-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e1e94-816">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="e1e94-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-817">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-817">Parameters</span></span>

| <span data-ttu-id="e1e94-818">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-818">Name</span></span> | <span data-ttu-id="e1e94-819">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-819">Type</span></span> | <span data-ttu-id="e1e94-820">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-820">Attributes</span></span> | <span data-ttu-id="e1e94-821">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e1e94-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e1e94-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e1e94-823">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="e1e94-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e1e94-824">Função</span><span class="sxs-lookup"><span data-stu-id="e1e94-824">Function</span></span> || <span data-ttu-id="e1e94-p140">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e1e94-828">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-828">Object</span></span> | <span data-ttu-id="e1e94-829">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-829">&lt;optional&gt;</span></span> | <span data-ttu-id="e1e94-830">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e1e94-831">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-831">Object</span></span> | <span data-ttu-id="e1e94-832">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-832">&lt;optional&gt;</span></span> | <span data-ttu-id="e1e94-833">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e1e94-834">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-834">function</span></span>| <span data-ttu-id="e1e94-835">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-835">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-836">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-837">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-837">Requirements</span></span>

|<span data-ttu-id="e1e94-838">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-838">Requirement</span></span>| <span data-ttu-id="e1e94-839">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-840">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e1e94-841">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-841">1.7</span></span> |
|[<span data-ttu-id="e1e94-842">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e1e94-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-843">ReadItem</span></span> |
|[<span data-ttu-id="e1e94-844">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e1e94-845">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="e1e94-846">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-846">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e1e94-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e1e94-848">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e1e94-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e1e94-852">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e1e94-853">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-854">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-854">Parameters</span></span>

|<span data-ttu-id="e1e94-855">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-855">Name</span></span>|<span data-ttu-id="e1e94-856">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-856">Type</span></span>|<span data-ttu-id="e1e94-857">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-857">Attributes</span></span>|<span data-ttu-id="e1e94-858">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="e1e94-859">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-859">String</span></span>||<span data-ttu-id="e1e94-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e1e94-862">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e1e94-862">String</span></span>||<span data-ttu-id="e1e94-863">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-863">The subject of the item to be attached.</span></span> <span data-ttu-id="e1e94-864">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e1e94-865">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-865">Object</span></span>|<span data-ttu-id="e1e94-866">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-866">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-867">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-868">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-868">Object</span></span>|<span data-ttu-id="e1e94-869">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-869">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-870">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-871">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-871">function</span></span>|<span data-ttu-id="e1e94-872">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-872">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-873">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e1e94-874">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e1e94-875">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="e1e94-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e1e94-876">Erros</span><span class="sxs-lookup"><span data-stu-id="e1e94-876">Errors</span></span>

|<span data-ttu-id="e1e94-877">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e1e94-877">Error code</span></span>|<span data-ttu-id="e1e94-878">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e1e94-879">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-880">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-880">Requirements</span></span>

|<span data-ttu-id="e1e94-881">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-881">Requirement</span></span>|<span data-ttu-id="e1e94-882">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-883">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-884">1.1</span><span class="sxs-lookup"><span data-stu-id="e1e94-884">1.1</span></span>|
|[<span data-ttu-id="e1e94-885">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-885">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-887">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-887">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-888">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-889">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-889">Example</span></span>

<span data-ttu-id="e1e94-890">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="e1e94-891">close()</span><span class="sxs-lookup"><span data-stu-id="e1e94-891">close()</span></span>

<span data-ttu-id="e1e94-892">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e1e94-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-895">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="e1e94-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e1e94-896">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="e1e94-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-897">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-897">Requirements</span></span>

|<span data-ttu-id="e1e94-898">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-898">Requirement</span></span>|<span data-ttu-id="e1e94-899">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-900">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-901">1.3</span><span class="sxs-lookup"><span data-stu-id="e1e94-901">1.3</span></span>|
|[<span data-ttu-id="e1e94-902">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-902">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-903">Restrito</span><span class="sxs-lookup"><span data-stu-id="e1e94-903">Restricted</span></span>|
|[<span data-ttu-id="e1e94-904">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-904">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-905">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-905">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="e1e94-906">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="e1e94-907">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-908">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-909">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="e1e94-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e1e94-910">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e1e94-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e1e94-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-914">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-914">Parameters</span></span>

|<span data-ttu-id="e1e94-915">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-915">Name</span></span>|<span data-ttu-id="e1e94-916">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-916">Type</span></span>|<span data-ttu-id="e1e94-917">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-917">Attributes</span></span>|<span data-ttu-id="e1e94-918">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e1e94-919">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-919">String &#124; Object</span></span>||<span data-ttu-id="e1e94-920">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="e1e94-920">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="e1e94-921">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e1e94-921">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e1e94-922">**OU**</span><span class="sxs-lookup"><span data-stu-id="e1e94-922">**OR**</span></span><br/><span data-ttu-id="e1e94-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e1e94-925">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-925">String</span></span>|<span data-ttu-id="e1e94-926">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-926">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e1e94-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e1e94-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-930">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-931">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e1e94-932">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-932">String</span></span>||<span data-ttu-id="e1e94-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e1e94-935">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-935">String</span></span>||<span data-ttu-id="e1e94-936">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="e1e94-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e1e94-937">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-937">String</span></span>||<span data-ttu-id="e1e94-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e1e94-940">Booliano</span><span class="sxs-lookup"><span data-stu-id="e1e94-940">Boolean</span></span>||<span data-ttu-id="e1e94-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e1e94-943">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-943">String</span></span>||<span data-ttu-id="e1e94-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e1e94-947">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-947">function</span></span>|<span data-ttu-id="e1e94-948">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-948">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-949">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-950">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-950">Requirements</span></span>

|<span data-ttu-id="e1e94-951">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-951">Requirement</span></span>|<span data-ttu-id="e1e94-952">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-953">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-954">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-954">1.0</span></span>|
|[<span data-ttu-id="e1e94-955">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-955">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-956">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-957">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-957">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-958">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e1e94-959">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e1e94-959">Examples</span></span>

<span data-ttu-id="e1e94-960">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e1e94-961">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="e1e94-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e1e94-962">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e1e94-963">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e1e94-964">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e1e94-965">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="e1e94-966">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="e1e94-967">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-968">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-969">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="e1e94-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e1e94-970">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e1e94-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e1e94-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-974">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-974">Parameters</span></span>

|<span data-ttu-id="e1e94-975">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-975">Name</span></span>|<span data-ttu-id="e1e94-976">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-976">Type</span></span>|<span data-ttu-id="e1e94-977">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-977">Attributes</span></span>|<span data-ttu-id="e1e94-978">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e1e94-979">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-979">String &#124; Object</span></span>||<span data-ttu-id="e1e94-980">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="e1e94-980">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="e1e94-981">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e1e94-981">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e1e94-982">**OU**</span><span class="sxs-lookup"><span data-stu-id="e1e94-982">**OR**</span></span><br/><span data-ttu-id="e1e94-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e1e94-985">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-985">String</span></span>|<span data-ttu-id="e1e94-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-986">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e1e94-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e1e94-990">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-990">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-991">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e1e94-992">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-992">String</span></span>||<span data-ttu-id="e1e94-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e1e94-995">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-995">String</span></span>||<span data-ttu-id="e1e94-996">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="e1e94-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e1e94-997">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-997">String</span></span>||<span data-ttu-id="e1e94-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e1e94-1000">Booliano</span><span class="sxs-lookup"><span data-stu-id="e1e94-1000">Boolean</span></span>||<span data-ttu-id="e1e94-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e1e94-1003">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1003">String</span></span>||<span data-ttu-id="e1e94-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1007">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1007">function</span></span>|<span data-ttu-id="e1e94-1008">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1009">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1010">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1010">Requirements</span></span>

|<span data-ttu-id="e1e94-1011">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1011">Requirement</span></span>|<span data-ttu-id="e1e94-1012">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1013">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1014">1.0</span></span>|
|[<span data-ttu-id="e1e94-1015">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1015">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1016">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1017">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1017">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1018">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e1e94-1019">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1019">Examples</span></span>

<span data-ttu-id="e1e94-1020">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e1e94-1021">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e1e94-1022">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e1e94-1023">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e1e94-1024">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e1e94-1025">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="e1e94-1026">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="e1e94-1027">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="e1e94-1028">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e1e94-1029">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="e1e94-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="e1e94-1030">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e1e94-1031">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1032">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1032">Parameters</span></span>

|<span data-ttu-id="e1e94-1033">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1033">Name</span></span>|<span data-ttu-id="e1e94-1034">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1034">Type</span></span>|<span data-ttu-id="e1e94-1035">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1035">Attributes</span></span>|<span data-ttu-id="e1e94-1036">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e1e94-1037">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1037">String</span></span>||<span data-ttu-id="e1e94-1038">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="e1e94-1039">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1039">Object</span></span>|<span data-ttu-id="e1e94-1040">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1041">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1042">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1042">Object</span></span>|<span data-ttu-id="e1e94-1043">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1044">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1045">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1045">function</span></span>|<span data-ttu-id="e1e94-1046">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1047">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1048">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1048">Requirements</span></span>

|<span data-ttu-id="e1e94-1049">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1049">Requirement</span></span>|<span data-ttu-id="e1e94-1050">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1051">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1052">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-1052">Preview</span></span>|
|[<span data-ttu-id="e1e94-1053">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1054">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1055">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1056">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1057">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1057">Returns:</span></span>

<span data-ttu-id="e1e94-1058">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1059">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1059">Example</span></span>

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
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e1e94-1060">getAttachmentsAsync ([opções], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e1e94-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e1e94-1061">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e1e94-1062">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1063">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1063">Parameters</span></span>

|<span data-ttu-id="e1e94-1064">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1064">Name</span></span>|<span data-ttu-id="e1e94-1065">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1065">Type</span></span>|<span data-ttu-id="e1e94-1066">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1066">Attributes</span></span>|<span data-ttu-id="e1e94-1067">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e1e94-1068">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1068">Object</span></span>|<span data-ttu-id="e1e94-1069">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1070">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1071">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1071">Object</span></span>|<span data-ttu-id="e1e94-1072">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1073">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1074">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1074">function</span></span>|<span data-ttu-id="e1e94-1075">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1076">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1077">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1077">Requirements</span></span>

|<span data-ttu-id="e1e94-1078">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1078">Requirement</span></span>|<span data-ttu-id="e1e94-1079">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1080">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1081">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-1081">Preview</span></span>|
|[<span data-ttu-id="e1e94-1082">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1082">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1083">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1084">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1084">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1085">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1086">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1086">Returns:</span></span>

<span data-ttu-id="e1e94-1087">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e1e94-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1088">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1088">Example</span></span>

<span data-ttu-id="e1e94-1089">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e1e94-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e1e94-1091">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1092">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-1093">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1093">Requirements</span></span>

|<span data-ttu-id="e1e94-1094">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1094">Requirement</span></span>|<span data-ttu-id="e1e94-1095">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1096">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1097">1.0</span></span>|
|[<span data-ttu-id="e1e94-1098">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1098">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1099">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1100">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1100">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1101">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1102">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1102">Returns:</span></span>

<span data-ttu-id="e1e94-1103">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1104">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1104">Example</span></span>

<span data-ttu-id="e1e94-1105">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e1e94-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e1e94-1107">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1108">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1109">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1109">Parameters</span></span>

|<span data-ttu-id="e1e94-1110">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1110">Name</span></span>|<span data-ttu-id="e1e94-1111">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1111">Type</span></span>|<span data-ttu-id="e1e94-1112">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="e1e94-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e1e94-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="e1e94-1114">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1115">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1115">Requirements</span></span>

|<span data-ttu-id="e1e94-1116">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1116">Requirement</span></span>|<span data-ttu-id="e1e94-1117">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1118">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1119">1.0</span></span>|
|[<span data-ttu-id="e1e94-1120">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1120">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1121">Restrito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1121">Restricted</span></span>|
|[<span data-ttu-id="e1e94-1122">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1122">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1123">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1124">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1124">Returns:</span></span>

<span data-ttu-id="e1e94-1125">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e1e94-1126">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e1e94-1127">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e1e94-1128">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="e1e94-1129">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="e1e94-1129">Value of `entityType`</span></span>|<span data-ttu-id="e1e94-1130">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="e1e94-1130">Type of objects in returned array</span></span>|<span data-ttu-id="e1e94-1131">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="e1e94-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="e1e94-1132">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1132">String</span></span>|<span data-ttu-id="e1e94-1133">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="e1e94-1134">Contato</span><span class="sxs-lookup"><span data-stu-id="e1e94-1134">Contact</span></span>|<span data-ttu-id="e1e94-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="e1e94-1136">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1136">String</span></span>|<span data-ttu-id="e1e94-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="e1e94-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e1e94-1138">MeetingSuggestion</span></span>|<span data-ttu-id="e1e94-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="e1e94-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e1e94-1140">PhoneNumber</span></span>|<span data-ttu-id="e1e94-1141">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="e1e94-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e1e94-1142">TaskSuggestion</span></span>|<span data-ttu-id="e1e94-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="e1e94-1144">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1144">String</span></span>|<span data-ttu-id="e1e94-1145">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e1e94-1145">**Restricted**</span></span>|

<span data-ttu-id="e1e94-1146">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e1e94-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1147">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1147">Example</span></span>

<span data-ttu-id="e1e94-1148">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e1e94-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e1e94-1150">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1151">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-1152">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1153">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1153">Parameters</span></span>

|<span data-ttu-id="e1e94-1154">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1154">Name</span></span>|<span data-ttu-id="e1e94-1155">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1155">Type</span></span>|<span data-ttu-id="e1e94-1156">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e1e94-1157">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1157">String</span></span>|<span data-ttu-id="e1e94-1158">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1159">Requirements</span></span>

|<span data-ttu-id="e1e94-1160">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1160">Requirement</span></span>|<span data-ttu-id="e1e94-1161">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1163">1.0</span></span>|
|[<span data-ttu-id="e1e94-1164">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1164">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1165">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1166">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1167">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1168">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1168">Returns:</span></span>

<span data-ttu-id="e1e94-p164">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e1e94-1171">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e1e94-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="e1e94-1172">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="e1e94-1173">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1173">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1174">Este método só tem suporte no Outlook 2016 ou posterior para Windows (versões clique para executar depois de 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1175">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1175">Parameters</span></span>
|<span data-ttu-id="e1e94-1176">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1176">Name</span></span>|<span data-ttu-id="e1e94-1177">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1177">Type</span></span>|<span data-ttu-id="e1e94-1178">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1178">Attributes</span></span>|<span data-ttu-id="e1e94-1179">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e1e94-1180">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-1180">Object</span></span>|<span data-ttu-id="e1e94-1181">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1182">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1183">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-1183">Object</span></span>|<span data-ttu-id="e1e94-1184">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1185">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1186">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1186">function</span></span>|<span data-ttu-id="e1e94-1187">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1188">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e1e94-1189">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="e1e94-1190">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1191">Requirements</span></span>

|<span data-ttu-id="e1e94-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1192">Requirement</span></span>|<span data-ttu-id="e1e94-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1195">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-1195">Preview</span></span>|
|[<span data-ttu-id="e1e94-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1197">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1199">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1200">Example</span></span>

```javascript
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

#### <a name="getregexmatches--object"></a><span data-ttu-id="e1e94-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e1e94-1202">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1203">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-p165">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e1e94-1207">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e1e94-1208">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e1e94-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-1212">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1212">Requirements</span></span>

|<span data-ttu-id="e1e94-1213">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1213">Requirement</span></span>|<span data-ttu-id="e1e94-1214">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1215">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1216">1.0</span></span>|
|[<span data-ttu-id="e1e94-1217">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1217">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1218">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1220">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1221">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1221">Returns:</span></span>

<span data-ttu-id="e1e94-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e1e94-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="e1e94-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e1e94-1225">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e1e94-1226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1226">Example</span></span>

<span data-ttu-id="e1e94-1227">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e1e94-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e1e94-1229">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1230">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-1231">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e1e94-p168">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1234">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1234">Parameters</span></span>

|<span data-ttu-id="e1e94-1235">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1235">Name</span></span>|<span data-ttu-id="e1e94-1236">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1236">Type</span></span>|<span data-ttu-id="e1e94-1237">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e1e94-1238">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1238">String</span></span>|<span data-ttu-id="e1e94-1239">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1240">Requirements</span></span>

|<span data-ttu-id="e1e94-1241">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1241">Requirement</span></span>|<span data-ttu-id="e1e94-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1244">1.0</span></span>|
|[<span data-ttu-id="e1e94-1245">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1246">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1247">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1248">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1249">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1249">Returns:</span></span>

<span data-ttu-id="e1e94-1250">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="e1e94-1251">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="e1e94-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e1e94-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e1e94-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e1e94-1253">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e1e94-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e1e94-1255">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e1e94-p169">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1258">Parameters</span></span>

|<span data-ttu-id="e1e94-1259">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1259">Name</span></span>|<span data-ttu-id="e1e94-1260">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1260">Type</span></span>|<span data-ttu-id="e1e94-1261">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1261">Attributes</span></span>|<span data-ttu-id="e1e94-1262">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="e1e94-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e1e94-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e1e94-p170">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="e1e94-1267">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1267">Object</span></span>|<span data-ttu-id="e1e94-1268">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1269">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1270">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1270">Object</span></span>|<span data-ttu-id="e1e94-1271">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1272">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1273">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1273">function</span></span>||<span data-ttu-id="e1e94-1274">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e1e94-1275">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e1e94-1276">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1277">Requirements</span></span>

|<span data-ttu-id="e1e94-1278">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1278">Requirement</span></span>|<span data-ttu-id="e1e94-1279">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="e1e94-1281">1.2</span></span>|
|[<span data-ttu-id="e1e94-1282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-1284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1285">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1286">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1286">Returns:</span></span>

<span data-ttu-id="e1e94-1287">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="e1e94-1288">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="e1e94-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e1e94-1289">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e1e94-1290">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e1e94-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e1e94-1292">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1292">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="e1e94-1293">As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1293">Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1294">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-1295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1295">Requirements</span></span>

|<span data-ttu-id="e1e94-1296">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1296">Requirement</span></span>|<span data-ttu-id="e1e94-1297">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="e1e94-1299">1.6</span></span>|
|[<span data-ttu-id="e1e94-1300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1301">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1302">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1303">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1304">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1304">Returns:</span></span>

<span data-ttu-id="e1e94-1305">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1306">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1306">Example</span></span>

<span data-ttu-id="e1e94-1307">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="e1e94-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e1e94-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="e1e94-p173">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e1e94-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1311">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e1e94-p174">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e1e94-1315">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e1e94-1316">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e1e94-p175">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1e94-1320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1320">Requirements</span></span>

|<span data-ttu-id="e1e94-1321">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1321">Requirement</span></span>|<span data-ttu-id="e1e94-1322">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1323">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="e1e94-1324">1.6</span></span>|
|[<span data-ttu-id="e1e94-1325">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1326">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1327">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1328">Read</span><span class="sxs-lookup"><span data-stu-id="e1e94-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e1e94-1329">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1329">Returns:</span></span>

<span data-ttu-id="e1e94-p176">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="e1e94-1332">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1332">Example</span></span>

<span data-ttu-id="e1e94-1333">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="e1e94-1334">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="e1e94-1335">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1336">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1336">Parameters</span></span>

|<span data-ttu-id="e1e94-1337">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1337">Name</span></span>|<span data-ttu-id="e1e94-1338">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1338">Type</span></span>|<span data-ttu-id="e1e94-1339">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1339">Attributes</span></span>|<span data-ttu-id="e1e94-1340">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e1e94-1341">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-1341">Object</span></span>|<span data-ttu-id="e1e94-1342">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1343">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1344">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1344">Object</span></span>|<span data-ttu-id="e1e94-1345">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1346">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1347">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1347">function</span></span>||<span data-ttu-id="e1e94-1348">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e1e94-1349">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e1e94-1350">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1351">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1351">Requirements</span></span>

|<span data-ttu-id="e1e94-1352">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1352">Requirement</span></span>|<span data-ttu-id="e1e94-1353">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1354">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1355">Visualização</span><span class="sxs-lookup"><span data-stu-id="e1e94-1355">Preview</span></span>|
|[<span data-ttu-id="e1e94-1356">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1356">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1357">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1358">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1358">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1359">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-1360">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e1e94-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e1e94-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e1e94-1362">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e1e94-p178">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1366">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1366">Parameters</span></span>

|<span data-ttu-id="e1e94-1367">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1367">Name</span></span>|<span data-ttu-id="e1e94-1368">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1368">Type</span></span>|<span data-ttu-id="e1e94-1369">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1369">Attributes</span></span>|<span data-ttu-id="e1e94-1370">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="e1e94-1371">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1371">function</span></span>||<span data-ttu-id="e1e94-1372">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e1e94-1373">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e1e94-1374">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="e1e94-1375">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1375">Object</span></span>|<span data-ttu-id="e1e94-1376">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1377">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e1e94-1378">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1379">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1379">Requirements</span></span>

|<span data-ttu-id="e1e94-1380">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1380">Requirement</span></span>|<span data-ttu-id="e1e94-1381">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1382">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="e1e94-1383">1.0</span></span>|
|[<span data-ttu-id="e1e94-1384">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1385">ReadItem</span></span>|
|[<span data-ttu-id="e1e94-1386">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1387">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-1388">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1388">Example</span></span>

<span data-ttu-id="e1e94-p181">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e1e94-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e1e94-1393">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e1e94-1394">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e1e94-1395">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="e1e94-1396">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e1e94-1397">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1398">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1398">Parameters</span></span>

|<span data-ttu-id="e1e94-1399">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1399">Name</span></span>|<span data-ttu-id="e1e94-1400">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1400">Type</span></span>|<span data-ttu-id="e1e94-1401">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1401">Attributes</span></span>|<span data-ttu-id="e1e94-1402">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e1e94-1403">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1403">String</span></span>||<span data-ttu-id="e1e94-1404">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="e1e94-1405">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1405">Object</span></span>|<span data-ttu-id="e1e94-1406">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1407">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1408">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1408">Object</span></span>|<span data-ttu-id="e1e94-1409">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1410">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1411">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1411">function</span></span>|<span data-ttu-id="e1e94-1412">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1413">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e1e94-1414">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e1e94-1415">Erros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1415">Errors</span></span>

|<span data-ttu-id="e1e94-1416">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e1e94-1416">Error code</span></span>|<span data-ttu-id="e1e94-1417">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="e1e94-1418">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1419">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1419">Requirements</span></span>

|<span data-ttu-id="e1e94-1420">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1420">Requirement</span></span>|<span data-ttu-id="e1e94-1421">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1422">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="e1e94-1423">1.1</span></span>|
|[<span data-ttu-id="e1e94-1424">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1424">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-1426">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1426">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1427">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-1428">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1428">Example</span></span>

<span data-ttu-id="e1e94-1429">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e1e94-1430">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e1e94-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e1e94-1431">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e1e94-1432">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1433">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1433">Parameters</span></span>

| <span data-ttu-id="e1e94-1434">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1434">Name</span></span> | <span data-ttu-id="e1e94-1435">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1435">Type</span></span> | <span data-ttu-id="e1e94-1436">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1436">Attributes</span></span> | <span data-ttu-id="e1e94-1437">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e1e94-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e1e94-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e1e94-1439">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e1e94-1440">Objeto</span><span class="sxs-lookup"><span data-stu-id="e1e94-1440">Object</span></span> | <span data-ttu-id="e1e94-1441">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="e1e94-1442">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e1e94-1443">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1443">Object</span></span> | <span data-ttu-id="e1e94-1444">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="e1e94-1445">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e1e94-1446">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1446">function</span></span>| <span data-ttu-id="e1e94-1447">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1448">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1449">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1449">Requirements</span></span>

|<span data-ttu-id="e1e94-1450">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1450">Requirement</span></span>| <span data-ttu-id="e1e94-1451">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1452">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e1e94-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="e1e94-1453">1.7</span></span> |
|[<span data-ttu-id="e1e94-1454">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1454">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e1e94-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1455">ReadItem</span></span> |
|[<span data-ttu-id="e1e94-1456">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e1e94-1456">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e1e94-1457">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1e94-1457">Compose or Read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="e1e94-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="e1e94-1459">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="e1e94-p183">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1463">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="e1e94-1464">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e1e94-p185">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e1e94-1468">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="e1e94-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e1e94-1469">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="e1e94-1470">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="e1e94-1471">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1472">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1472">Parameters</span></span>

|<span data-ttu-id="e1e94-1473">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1473">Name</span></span>|<span data-ttu-id="e1e94-1474">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1474">Type</span></span>|<span data-ttu-id="e1e94-1475">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1475">Attributes</span></span>|<span data-ttu-id="e1e94-1476">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e1e94-1477">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1477">Object</span></span>|<span data-ttu-id="e1e94-1478">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1479">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1480">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1480">Object</span></span>|<span data-ttu-id="e1e94-1481">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1482">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1483">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1483">function</span></span>||<span data-ttu-id="e1e94-1484">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e1e94-1485">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1486">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1486">Requirements</span></span>

|<span data-ttu-id="e1e94-1487">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1487">Requirement</span></span>|<span data-ttu-id="e1e94-1488">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1489">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="e1e94-1490">1.3</span></span>|
|[<span data-ttu-id="e1e94-1491">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-1493">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1494">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e1e94-1495">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e1e94-p187">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e1e94-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e1e94-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e1e94-1499">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e1e94-p188">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e1e94-1503">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e1e94-1503">Parameters</span></span>

|<span data-ttu-id="e1e94-1504">Name</span><span class="sxs-lookup"><span data-stu-id="e1e94-1504">Name</span></span>|<span data-ttu-id="e1e94-1505">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1505">Type</span></span>|<span data-ttu-id="e1e94-1506">Atributos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1506">Attributes</span></span>|<span data-ttu-id="e1e94-1507">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e94-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="e1e94-1508">String</span><span class="sxs-lookup"><span data-stu-id="e1e94-1508">String</span></span>||<span data-ttu-id="e1e94-p189">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="e1e94-1512">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1512">Object</span></span>|<span data-ttu-id="e1e94-1513">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1514">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e1e94-1515">Object</span><span class="sxs-lookup"><span data-stu-id="e1e94-1515">Object</span></span>|<span data-ttu-id="e1e94-1516">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-1517">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e1e94-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e1e94-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e1e94-1519">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e1e94-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="e1e94-p190">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e1e94-p191">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="e1e94-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e1e94-1524">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="e1e94-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="e1e94-1525">function</span><span class="sxs-lookup"><span data-stu-id="e1e94-1525">function</span></span>||<span data-ttu-id="e1e94-1526">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e1e94-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e1e94-1527">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e94-1527">Requirements</span></span>

|<span data-ttu-id="e1e94-1528">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1e94-1528">Requirement</span></span>|<span data-ttu-id="e1e94-1529">Valor</span><span class="sxs-lookup"><span data-stu-id="e1e94-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1e94-1530">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1e94-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e1e94-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="e1e94-1531">1.2</span></span>|
|[<span data-ttu-id="e1e94-1532">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e1e94-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e1e94-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="e1e94-1534">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1e94-1534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e1e94-1535">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1e94-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e1e94-1536">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1e94-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
