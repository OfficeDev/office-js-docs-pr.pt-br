---
title: Office. Context. Mailbox. Item-visualização do conjunto de requisitos
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: d72d7acc285b1a5cf371b1c5e6b2a0a1653d2091
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952240"
---
# <a name="item"></a><span data-ttu-id="37ea2-102">item</span><span class="sxs-lookup"><span data-stu-id="37ea2-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="37ea2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="37ea2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="37ea2-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="37ea2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-106">Requirements</span></span>

|<span data-ttu-id="37ea2-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-107">Requirement</span></span>|<span data-ttu-id="37ea2-108">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-110">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-110">1.0</span></span>|
|[<span data-ttu-id="37ea2-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="37ea2-112">Restricted</span></span>|
|[<span data-ttu-id="37ea2-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="37ea2-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="37ea2-115">Members and methods</span></span>

| <span data-ttu-id="37ea2-116">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-116">Member</span></span> | <span data-ttu-id="37ea2-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="37ea2-118">attachments</span><span class="sxs-lookup"><span data-stu-id="37ea2-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="37ea2-119">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-119">Member</span></span> |
| [<span data-ttu-id="37ea2-120">bcc</span><span class="sxs-lookup"><span data-stu-id="37ea2-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="37ea2-121">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-121">Member</span></span> |
| [<span data-ttu-id="37ea2-122">body</span><span class="sxs-lookup"><span data-stu-id="37ea2-122">body</span></span>](#body-body) | <span data-ttu-id="37ea2-123">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-123">Member</span></span> |
| [<span data-ttu-id="37ea2-124">Categorias</span><span class="sxs-lookup"><span data-stu-id="37ea2-124">categories</span></span>](#categories-categories) | <span data-ttu-id="37ea2-125">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-125">Member</span></span> |
| [<span data-ttu-id="37ea2-126">cc</span><span class="sxs-lookup"><span data-stu-id="37ea2-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="37ea2-127">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-127">Member</span></span> |
| [<span data-ttu-id="37ea2-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="37ea2-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="37ea2-129">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-129">Member</span></span> |
| [<span data-ttu-id="37ea2-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="37ea2-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="37ea2-131">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-131">Member</span></span> |
| [<span data-ttu-id="37ea2-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="37ea2-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="37ea2-133">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-133">Member</span></span> |
| [<span data-ttu-id="37ea2-134">end</span><span class="sxs-lookup"><span data-stu-id="37ea2-134">end</span></span>](#end-datetime) | <span data-ttu-id="37ea2-135">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-135">Member</span></span> |
| [<span data-ttu-id="37ea2-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="37ea2-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="37ea2-137">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-137">Member</span></span> |
| [<span data-ttu-id="37ea2-138">from</span><span class="sxs-lookup"><span data-stu-id="37ea2-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="37ea2-139">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-139">Member</span></span> |
| [<span data-ttu-id="37ea2-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="37ea2-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="37ea2-141">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-141">Member</span></span> |
| [<span data-ttu-id="37ea2-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="37ea2-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="37ea2-143">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-143">Member</span></span> |
| [<span data-ttu-id="37ea2-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="37ea2-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="37ea2-145">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-145">Member</span></span> |
| [<span data-ttu-id="37ea2-146">itemId</span><span class="sxs-lookup"><span data-stu-id="37ea2-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="37ea2-147">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-147">Member</span></span> |
| [<span data-ttu-id="37ea2-148">itemType</span><span class="sxs-lookup"><span data-stu-id="37ea2-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="37ea2-149">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-149">Member</span></span> |
| [<span data-ttu-id="37ea2-150">location</span><span class="sxs-lookup"><span data-stu-id="37ea2-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="37ea2-151">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-151">Member</span></span> |
| [<span data-ttu-id="37ea2-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="37ea2-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="37ea2-153">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-153">Member</span></span> |
| [<span data-ttu-id="37ea2-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="37ea2-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="37ea2-155">Member</span><span class="sxs-lookup"><span data-stu-id="37ea2-155">Member</span></span> |
| [<span data-ttu-id="37ea2-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="37ea2-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="37ea2-157">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-157">Member</span></span> |
| [<span data-ttu-id="37ea2-158">organizer</span><span class="sxs-lookup"><span data-stu-id="37ea2-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="37ea2-159">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-159">Member</span></span> |
| [<span data-ttu-id="37ea2-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="37ea2-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="37ea2-161">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-161">Member</span></span> |
| [<span data-ttu-id="37ea2-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="37ea2-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="37ea2-163">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-163">Member</span></span> |
| [<span data-ttu-id="37ea2-164">sender</span><span class="sxs-lookup"><span data-stu-id="37ea2-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="37ea2-165">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-165">Member</span></span> |
| [<span data-ttu-id="37ea2-166">seriesid</span><span class="sxs-lookup"><span data-stu-id="37ea2-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="37ea2-167">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-167">Member</span></span> |
| [<span data-ttu-id="37ea2-168">start</span><span class="sxs-lookup"><span data-stu-id="37ea2-168">start</span></span>](#start-datetime) | <span data-ttu-id="37ea2-169">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-169">Member</span></span> |
| [<span data-ttu-id="37ea2-170">subject</span><span class="sxs-lookup"><span data-stu-id="37ea2-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="37ea2-171">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-171">Member</span></span> |
| [<span data-ttu-id="37ea2-172">to</span><span class="sxs-lookup"><span data-stu-id="37ea2-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="37ea2-173">Membro</span><span class="sxs-lookup"><span data-stu-id="37ea2-173">Member</span></span> |
| [<span data-ttu-id="37ea2-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="37ea2-175">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-175">Method</span></span> |
| [<span data-ttu-id="37ea2-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="37ea2-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="37ea2-177">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-177">Method</span></span> |
| [<span data-ttu-id="37ea2-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="37ea2-179">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-179">Method</span></span> |
| [<span data-ttu-id="37ea2-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="37ea2-181">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-181">Method</span></span> |
| [<span data-ttu-id="37ea2-182">close</span><span class="sxs-lookup"><span data-stu-id="37ea2-182">close</span></span>](#close) | <span data-ttu-id="37ea2-183">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-183">Method</span></span> |
| [<span data-ttu-id="37ea2-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="37ea2-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="37ea2-185">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-185">Method</span></span> |
| [<span data-ttu-id="37ea2-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="37ea2-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="37ea2-187">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-187">Method</span></span> |
| [<span data-ttu-id="37ea2-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="37ea2-189">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-189">Method</span></span> |
| [<span data-ttu-id="37ea2-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="37ea2-191">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-191">Method</span></span> |
| [<span data-ttu-id="37ea2-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="37ea2-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="37ea2-193">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-193">Method</span></span> |
| [<span data-ttu-id="37ea2-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="37ea2-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="37ea2-195">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-195">Method</span></span> |
| [<span data-ttu-id="37ea2-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="37ea2-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="37ea2-197">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-197">Method</span></span> |
| [<span data-ttu-id="37ea2-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="37ea2-199">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-199">Method</span></span> |
| [<span data-ttu-id="37ea2-200">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="37ea2-200">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="37ea2-201">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-201">Method</span></span> |
| [<span data-ttu-id="37ea2-202">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="37ea2-202">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="37ea2-203">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-203">Method</span></span> |
| [<span data-ttu-id="37ea2-204">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-204">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="37ea2-205">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-205">Method</span></span> |
| [<span data-ttu-id="37ea2-206">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="37ea2-206">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="37ea2-207">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-207">Method</span></span> |
| [<span data-ttu-id="37ea2-208">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="37ea2-208">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="37ea2-209">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-209">Method</span></span> |
| [<span data-ttu-id="37ea2-210">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-210">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="37ea2-211">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-211">Method</span></span> |
| [<span data-ttu-id="37ea2-212">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-212">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="37ea2-213">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-213">Method</span></span> |
| [<span data-ttu-id="37ea2-214">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-214">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="37ea2-215">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-215">Method</span></span> |
| [<span data-ttu-id="37ea2-216">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-216">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="37ea2-217">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-217">Method</span></span> |
| [<span data-ttu-id="37ea2-218">saveAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-218">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="37ea2-219">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-219">Method</span></span> |
| [<span data-ttu-id="37ea2-220">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="37ea2-220">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="37ea2-221">Método</span><span class="sxs-lookup"><span data-stu-id="37ea2-221">Method</span></span> |

### <a name="example"></a><span data-ttu-id="37ea2-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-222">Example</span></span>

<span data-ttu-id="37ea2-223">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="37ea2-223">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="37ea2-224">Membros</span><span class="sxs-lookup"><span data-stu-id="37ea2-224">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="37ea2-225">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37ea2-225">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="37ea2-226">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="37ea2-226">Gets the item's attachments as an array.</span></span> <span data-ttu-id="37ea2-227">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-227">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-228">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="37ea2-228">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="37ea2-229">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="37ea2-229">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-230">Type</span></span>

*   <span data-ttu-id="37ea2-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37ea2-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-232">Requirements</span></span>

|<span data-ttu-id="37ea2-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-233">Requirement</span></span>|<span data-ttu-id="37ea2-234">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-236">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-236">1.0</span></span>|
|[<span data-ttu-id="37ea2-237">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-238">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-240">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-240">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-241">Example</span></span>

<span data-ttu-id="37ea2-242">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-242">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="37ea2-243">CCO: [destinatários](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-243">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="37ea2-244">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-244">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="37ea2-245">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="37ea2-245">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-246">Type</span></span>

*   [<span data-ttu-id="37ea2-247">Destinatários</span><span class="sxs-lookup"><span data-stu-id="37ea2-247">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="37ea2-248">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-248">Requirements</span></span>

|<span data-ttu-id="37ea2-249">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-249">Requirement</span></span>|<span data-ttu-id="37ea2-250">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-251">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-252">1.1</span><span class="sxs-lookup"><span data-stu-id="37ea2-252">1.1</span></span>|
|[<span data-ttu-id="37ea2-253">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-254">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-255">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-256">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-256">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-257">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-257">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="37ea2-258">corpo: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="37ea2-258">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="37ea2-259">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-259">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-260">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-260">Type</span></span>

*   [<span data-ttu-id="37ea2-261">Body</span><span class="sxs-lookup"><span data-stu-id="37ea2-261">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="37ea2-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-262">Requirements</span></span>

|<span data-ttu-id="37ea2-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-263">Requirement</span></span>|<span data-ttu-id="37ea2-264">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-265">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-266">1.1</span><span class="sxs-lookup"><span data-stu-id="37ea2-266">1.1</span></span>|
|[<span data-ttu-id="37ea2-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-268">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-269">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-270">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-270">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-271">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-271">Example</span></span>

<span data-ttu-id="37ea2-272">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="37ea2-272">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="37ea2-273">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-273">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="37ea2-274">Categorias: [categorias](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="37ea2-274">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="37ea2-275">Obtém um objeto que fornece métodos para gerenciar as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-275">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-276">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-276">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-277">Type</span></span>

*   [<span data-ttu-id="37ea2-278">Categories</span><span class="sxs-lookup"><span data-stu-id="37ea2-278">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="37ea2-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-279">Requirements</span></span>

|<span data-ttu-id="37ea2-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-280">Requirement</span></span>|<span data-ttu-id="37ea2-281">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-283">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-283">Preview</span></span>|
|[<span data-ttu-id="37ea2-284">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-285">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-286">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-287">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-288">Example</span></span>

<span data-ttu-id="37ea2-289">Este exemplo obtém as categorias do item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-289">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="37ea2-290">[destinatários](/javascript/api/outlook/office.recipients) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="37ea2-290">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="37ea2-291">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-291">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="37ea2-292">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-292">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-293">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-293">Read mode</span></span>

<span data-ttu-id="37ea2-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-296">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-296">Compose mode</span></span>

<span data-ttu-id="37ea2-297">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-297">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-298">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-298">Type</span></span>

*   <span data-ttu-id="37ea2-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-300">Requirements</span></span>

|<span data-ttu-id="37ea2-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-301">Requirement</span></span>|<span data-ttu-id="37ea2-302">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-304">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-304">1.0</span></span>|
|[<span data-ttu-id="37ea2-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-306">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-307">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-308">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-308">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="37ea2-309">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="37ea2-309">(nullable) conversationId: String</span></span>

<span data-ttu-id="37ea2-310">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="37ea2-310">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="37ea2-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="37ea2-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-315">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-315">Type</span></span>

*   <span data-ttu-id="37ea2-316">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-316">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-317">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-317">Requirements</span></span>

|<span data-ttu-id="37ea2-318">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-318">Requirement</span></span>|<span data-ttu-id="37ea2-319">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-320">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-321">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-321">1.0</span></span>|
|[<span data-ttu-id="37ea2-322">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-323">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-324">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-325">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-325">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-326">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-326">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="37ea2-327">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="37ea2-327">dateTimeCreated: Date</span></span>

<span data-ttu-id="37ea2-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-330">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-330">Type</span></span>

*   <span data-ttu-id="37ea2-331">Data</span><span class="sxs-lookup"><span data-stu-id="37ea2-331">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-332">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-332">Requirements</span></span>

|<span data-ttu-id="37ea2-333">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-333">Requirement</span></span>|<span data-ttu-id="37ea2-334">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-335">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-336">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-336">1.0</span></span>|
|[<span data-ttu-id="37ea2-337">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-338">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-339">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-340">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-340">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-341">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-341">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="37ea2-342">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="37ea2-342">dateTimeModified: Date</span></span>

<span data-ttu-id="37ea2-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-345">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-345">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-346">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-346">Type</span></span>

*   <span data-ttu-id="37ea2-347">Data</span><span class="sxs-lookup"><span data-stu-id="37ea2-347">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-348">Requirements</span></span>

|<span data-ttu-id="37ea2-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-349">Requirement</span></span>|<span data-ttu-id="37ea2-350">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-352">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-352">1.0</span></span>|
|[<span data-ttu-id="37ea2-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-354">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-356">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-357">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-357">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="37ea2-358">fim: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="37ea2-358">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="37ea2-359">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="37ea2-359">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="37ea2-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-362">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-362">Read mode</span></span>

<span data-ttu-id="37ea2-363">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-363">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-364">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-364">Compose mode</span></span>

<span data-ttu-id="37ea2-365">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-365">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="37ea2-366">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="37ea2-366">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="37ea2-367">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-367">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="37ea2-368">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-368">Type</span></span>

*   <span data-ttu-id="37ea2-369">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="37ea2-369">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-370">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-370">Requirements</span></span>

|<span data-ttu-id="37ea2-371">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-371">Requirement</span></span>|<span data-ttu-id="37ea2-372">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-372">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-373">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-373">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-374">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-374">1.0</span></span>|
|[<span data-ttu-id="37ea2-375">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-375">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-376">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-376">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-377">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-377">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-378">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-378">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="37ea2-379">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="37ea2-379">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="37ea2-380">Obtém ou define os locais de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-380">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-381">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-381">Read mode</span></span>

<span data-ttu-id="37ea2-382">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que permite que você obtenha o conjunto de locais (cada um representado por um objeto [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associado ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-382">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-383">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-383">Compose mode</span></span>

<span data-ttu-id="37ea2-384">A `enhancedLocation` propriedade retorna um objeto [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) que fornece métodos para obter, remover ou adicionar locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-385">Type</span></span>

*   [<span data-ttu-id="37ea2-386">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="37ea2-386">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="37ea2-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-387">Requirements</span></span>

|<span data-ttu-id="37ea2-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-388">Requirement</span></span>|<span data-ttu-id="37ea2-389">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-391">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-391">Preview</span></span>|
|[<span data-ttu-id="37ea2-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-393">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-394">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-395">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-395">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-396">Example</span></span>

<span data-ttu-id="37ea2-397">O exemplo a seguir obtém os locais atuais associados ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-397">The following example gets the current locations associated with the appointment.</span></span>

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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="37ea2-398">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="37ea2-398">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="37ea2-399">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-399">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="37ea2-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-402">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-402">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-403">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-403">Read mode</span></span>

<span data-ttu-id="37ea2-404">A `from` propriedade retorna um `EmailAddressDetails` objeto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-404">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-405">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-405">Compose mode</span></span>

<span data-ttu-id="37ea2-406">A `from` propriedade retorna um `From` objeto que fornece um método para obter o valor de.</span><span class="sxs-lookup"><span data-stu-id="37ea2-406">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-407">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-407">Type</span></span>

*   <span data-ttu-id="37ea2-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="37ea2-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-409">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-409">Requirements</span></span>

|<span data-ttu-id="37ea2-410">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-410">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="37ea2-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-412">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-412">1.0</span></span>|<span data-ttu-id="37ea2-413">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-413">1.7</span></span>|
|[<span data-ttu-id="37ea2-414">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-415">ReadItem</span></span>|<span data-ttu-id="37ea2-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-416">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-418">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-418">Read</span></span>|<span data-ttu-id="37ea2-419">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-419">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="37ea2-420">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="37ea2-420">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="37ea2-421">Obtém ou define os cabeçalhos de Internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-421">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-422">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-422">Type</span></span>

*   [<span data-ttu-id="37ea2-423">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="37ea2-423">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="37ea2-424">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-424">Requirements</span></span>

|<span data-ttu-id="37ea2-425">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-425">Requirement</span></span>|<span data-ttu-id="37ea2-426">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-427">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-428">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-428">Preview</span></span>|
|[<span data-ttu-id="37ea2-429">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-430">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-431">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-432">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-432">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-433">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-433">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="37ea2-434">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-434">internetMessageId: String</span></span>

<span data-ttu-id="37ea2-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-437">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-437">Type</span></span>

*   <span data-ttu-id="37ea2-438">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-438">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-439">Requirements</span></span>

|<span data-ttu-id="37ea2-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-440">Requirement</span></span>|<span data-ttu-id="37ea2-441">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-443">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-443">1.0</span></span>|
|[<span data-ttu-id="37ea2-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-445">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-446">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-447">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-447">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-448">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="37ea2-449">doclass: String</span><span class="sxs-lookup"><span data-stu-id="37ea2-449">itemClass: String</span></span>

<span data-ttu-id="37ea2-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="37ea2-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="37ea2-454">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-454">Type</span></span>|<span data-ttu-id="37ea2-455">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-455">Description</span></span>|<span data-ttu-id="37ea2-456">classe de item</span><span class="sxs-lookup"><span data-stu-id="37ea2-456">item class</span></span>|
|---|---|---|
|<span data-ttu-id="37ea2-457">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="37ea2-457">Appointment items</span></span>|<span data-ttu-id="37ea2-458">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-458">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="37ea2-459">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="37ea2-459">Message items</span></span>|<span data-ttu-id="37ea2-460">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="37ea2-460">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="37ea2-461">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-461">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-462">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-462">Type</span></span>

*   <span data-ttu-id="37ea2-463">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-464">Requirements</span></span>

|<span data-ttu-id="37ea2-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-465">Requirement</span></span>|<span data-ttu-id="37ea2-466">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-468">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-468">1.0</span></span>|
|[<span data-ttu-id="37ea2-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-470">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-472">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-473">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="37ea2-474">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="37ea2-474">(nullable) itemId: String</span></span>

<span data-ttu-id="37ea2-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-477">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="37ea2-477">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="37ea2-478">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="37ea2-478">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="37ea2-479">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="37ea2-479">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="37ea2-480">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="37ea2-480">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="37ea2-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-483">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-483">Type</span></span>

*   <span data-ttu-id="37ea2-484">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-484">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-485">Requirements</span></span>

|<span data-ttu-id="37ea2-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-486">Requirement</span></span>|<span data-ttu-id="37ea2-487">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-488">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-489">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-489">1.0</span></span>|
|[<span data-ttu-id="37ea2-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-491">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-493">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-493">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-494">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-494">Example</span></span>

<span data-ttu-id="37ea2-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="37ea2-497">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="37ea2-497">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="37ea2-498">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="37ea2-498">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="37ea2-499">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-499">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-500">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-500">Type</span></span>

*   [<span data-ttu-id="37ea2-501">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="37ea2-501">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="37ea2-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-502">Requirements</span></span>

|<span data-ttu-id="37ea2-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-503">Requirement</span></span>|<span data-ttu-id="37ea2-504">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-506">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-506">1.0</span></span>|
|[<span data-ttu-id="37ea2-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-508">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-509">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-510">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-510">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-511">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-511">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="37ea2-512">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="37ea2-512">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="37ea2-513">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-513">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-514">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-514">Read mode</span></span>

<span data-ttu-id="37ea2-515">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-515">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-516">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-516">Compose mode</span></span>

<span data-ttu-id="37ea2-517">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-517">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-518">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-518">Type</span></span>

*   <span data-ttu-id="37ea2-519">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="37ea2-519">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-520">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-520">Requirements</span></span>

|<span data-ttu-id="37ea2-521">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-521">Requirement</span></span>|<span data-ttu-id="37ea2-522">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-522">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-523">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-524">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-524">1.0</span></span>|
|[<span data-ttu-id="37ea2-525">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-526">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-527">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-528">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-528">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="37ea2-529">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-529">normalizedSubject: String</span></span>

<span data-ttu-id="37ea2-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="37ea2-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="37ea2-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-534">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-534">Type</span></span>

*   <span data-ttu-id="37ea2-535">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-535">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-536">Requirements</span></span>

|<span data-ttu-id="37ea2-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-537">Requirement</span></span>|<span data-ttu-id="37ea2-538">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-540">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-540">1.0</span></span>|
|[<span data-ttu-id="37ea2-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-542">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-544">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-545">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="37ea2-546">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="37ea2-546">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="37ea2-547">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-547">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-548">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-548">Type</span></span>

*   [<span data-ttu-id="37ea2-549">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="37ea2-549">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="37ea2-550">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-550">Requirements</span></span>

|<span data-ttu-id="37ea2-551">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-551">Requirement</span></span>|<span data-ttu-id="37ea2-552">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-553">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-554">1.3</span><span class="sxs-lookup"><span data-stu-id="37ea2-554">1.3</span></span>|
|[<span data-ttu-id="37ea2-555">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-556">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-557">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-558">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-558">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-559">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-559">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="37ea2-560">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz</span><span class="sxs-lookup"><span data-stu-id="37ea2-560">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="37ea2-561">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="37ea2-561">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="37ea2-562">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-562">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-563">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-563">Read mode</span></span>

<span data-ttu-id="37ea2-564">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-564">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-565">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-565">Compose mode</span></span>

<span data-ttu-id="37ea2-566">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-566">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-567">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-567">Type</span></span>

*   <span data-ttu-id="37ea2-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-569">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-569">Requirements</span></span>

|<span data-ttu-id="37ea2-570">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-570">Requirement</span></span>|<span data-ttu-id="37ea2-571">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-571">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-572">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-572">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-573">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-573">1.0</span></span>|
|[<span data-ttu-id="37ea2-574">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-574">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-575">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-575">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-576">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-577">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-577">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="37ea2-578">organizador: [](/javascript/api/outlook/office.emailaddressdetails)|[organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37ea2-578">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="37ea2-579">Obtém o endereço de email do organizador de uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-579">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-580">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-580">Read mode</span></span>

<span data-ttu-id="37ea2-581">A `organizer` propriedade retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-581">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-582">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-582">Compose mode</span></span>

<span data-ttu-id="37ea2-583">A `organizer` propriedade retorna um [](/javascript/api/outlook/office.organizer) objeto organizador que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="37ea2-583">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="37ea2-584">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-584">Type</span></span>

*   <span data-ttu-id="37ea2-585">[](/javascript/api/outlook/office.emailaddressdetails) | [Organizador](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37ea2-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-586">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-586">Requirements</span></span>

|<span data-ttu-id="37ea2-587">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-587">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="37ea2-588">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-589">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-589">1.0</span></span>|<span data-ttu-id="37ea2-590">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-590">1.7</span></span>|
|[<span data-ttu-id="37ea2-591">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-591">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-592">ReadItem</span></span>|<span data-ttu-id="37ea2-593">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-593">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-594">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-595">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-595">Read</span></span>|<span data-ttu-id="37ea2-596">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-596">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="37ea2-597">(anulável) recorrência [](/javascript/api/outlook/office.recurrence) : recorrência</span><span class="sxs-lookup"><span data-stu-id="37ea2-597">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="37ea2-598">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-598">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="37ea2-599">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-599">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="37ea2-600">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-600">Read and compose modes for appointment items.</span></span> <span data-ttu-id="37ea2-601">Modo de leitura para itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-601">Read mode for meeting request items.</span></span>

<span data-ttu-id="37ea2-602">A `recurrence` propriedade retorna um [](/javascript/api/outlook/office.recurrence) objeto de recorrência para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="37ea2-602">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="37ea2-603">`null`é retornado para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-603">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="37ea2-604">`undefined`é retornado para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-604">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="37ea2-605">Observação: as solicitações de reunião `itemClass` têm um valor IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="37ea2-605">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="37ea2-606">Observação: se o objeto Recurrence é `null`, isso indica que o objeto é um único compromisso ou uma solicitação de reunião de um único compromisso e não uma parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="37ea2-606">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-607">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-607">Read mode</span></span>

<span data-ttu-id="37ea2-608">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que representa a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-608">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="37ea2-609">Isso está disponível para compromissos e solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-609">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-610">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-610">Compose mode</span></span>

<span data-ttu-id="37ea2-611">A `recurrence` propriedade retorna um objeto [Recurrence](/javascript/api/outlook/office.recurrence) que fornece métodos para gerenciar a recorrência do compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-611">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="37ea2-612">Isso está disponível para compromissos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-612">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="37ea2-613">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-613">Type</span></span>

* [<span data-ttu-id="37ea2-614">Recorrência</span><span class="sxs-lookup"><span data-stu-id="37ea2-614">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="37ea2-615">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-615">Requirement</span></span>|<span data-ttu-id="37ea2-616">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-617">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-618">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-618">1.7</span></span>|
|[<span data-ttu-id="37ea2-619">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-620">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-621">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-622">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-622">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="37ea2-623">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[](/javascript/api/outlook/office.recipients) de matriz</span><span class="sxs-lookup"><span data-stu-id="37ea2-623">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="37ea2-624">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="37ea2-624">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="37ea2-625">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-625">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-626">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-626">Read mode</span></span>

<span data-ttu-id="37ea2-627">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-627">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-628">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-628">Compose mode</span></span>

<span data-ttu-id="37ea2-629">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-629">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-630">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-630">Type</span></span>

*   <span data-ttu-id="37ea2-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-632">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-632">Requirements</span></span>

|<span data-ttu-id="37ea2-633">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-633">Requirement</span></span>|<span data-ttu-id="37ea2-634">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-635">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-636">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-636">1.0</span></span>|
|[<span data-ttu-id="37ea2-637">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-638">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-639">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-640">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-640">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="37ea2-641">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="37ea2-641">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="37ea2-p128">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="37ea2-p129">As propriedades [`from`](#from-emailaddressdetailsfrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-646">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-646">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-647">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-647">Type</span></span>

*   [<span data-ttu-id="37ea2-648">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37ea2-648">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="37ea2-649">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-649">Requirements</span></span>

|<span data-ttu-id="37ea2-650">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-650">Requirement</span></span>|<span data-ttu-id="37ea2-651">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-651">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-652">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-653">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-653">1.0</span></span>|
|[<span data-ttu-id="37ea2-654">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-655">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-656">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-657">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-657">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-658">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-658">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="37ea2-659">(Nullable) seriesid: String</span><span class="sxs-lookup"><span data-stu-id="37ea2-659">(nullable) seriesId: String</span></span>

<span data-ttu-id="37ea2-660">Obtém a ID da série à qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="37ea2-660">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="37ea2-661">No OWA e no Outlook, `seriesId` o retorna a ID dos serviços Web do Exchange (EWS) do item pai (série) ao qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="37ea2-661">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="37ea2-662">No entanto, no iOS e no `seriesId` Android, o retorna a ID do REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="37ea2-662">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-663">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="37ea2-663">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="37ea2-664">A `seriesId` propriedade não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="37ea2-664">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="37ea2-665">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="37ea2-665">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="37ea2-666">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="37ea2-666">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="37ea2-667">A `seriesId` propriedade retorna `null` para itens que não têm itens pai, como compromissos únicos, itens de série ou solicitações de reunião e retornam `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="37ea2-667">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="37ea2-668">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-668">Type</span></span>

* <span data-ttu-id="37ea2-669">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-669">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-670">Requirements</span></span>

|<span data-ttu-id="37ea2-671">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-671">Requirement</span></span>|<span data-ttu-id="37ea2-672">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-673">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-674">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-674">1.7</span></span>|
|[<span data-ttu-id="37ea2-675">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-676">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-677">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-678">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-678">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-679">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-679">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="37ea2-680">Início: data | [Tempo](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="37ea2-680">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="37ea2-681">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="37ea2-681">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="37ea2-p132">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-684">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-684">Read mode</span></span>

<span data-ttu-id="37ea2-685">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-685">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-686">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-686">Compose mode</span></span>

<span data-ttu-id="37ea2-687">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-687">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="37ea2-688">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="37ea2-688">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="37ea2-689">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-689">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="37ea2-690">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-690">Type</span></span>

*   <span data-ttu-id="37ea2-691">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="37ea2-691">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-692">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-692">Requirements</span></span>

|<span data-ttu-id="37ea2-693">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-693">Requirement</span></span>|<span data-ttu-id="37ea2-694">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-695">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-696">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-696">1.0</span></span>|
|[<span data-ttu-id="37ea2-697">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-698">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-699">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-700">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-700">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="37ea2-701">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="37ea2-701">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="37ea2-702">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-702">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="37ea2-703">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="37ea2-703">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-704">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-704">Read mode</span></span>

<span data-ttu-id="37ea2-p133">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="37ea2-707">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="37ea2-707">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-708">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-708">Compose mode</span></span>
<span data-ttu-id="37ea2-709">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-709">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-710">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-710">Type</span></span>

*   <span data-ttu-id="37ea2-711">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="37ea2-711">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-712">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-712">Requirements</span></span>

|<span data-ttu-id="37ea2-713">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-713">Requirement</span></span>|<span data-ttu-id="37ea2-714">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-715">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-716">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-716">1.0</span></span>|
|[<span data-ttu-id="37ea2-717">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-718">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-719">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-720">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-720">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="37ea2-721">para: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-721">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="37ea2-722">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-722">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="37ea2-723">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-723">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37ea2-724">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="37ea2-724">Read mode</span></span>

<span data-ttu-id="37ea2-p135">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="37ea2-727">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="37ea2-727">Compose mode</span></span>

<span data-ttu-id="37ea2-728">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-728">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37ea2-729">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-729">Type</span></span>

*   <span data-ttu-id="37ea2-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37ea2-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-731">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-731">Requirements</span></span>

|<span data-ttu-id="37ea2-732">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-732">Requirement</span></span>|<span data-ttu-id="37ea2-733">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-734">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-735">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-735">1.0</span></span>|
|[<span data-ttu-id="37ea2-736">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-737">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-738">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-739">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-739">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="37ea2-740">Métodos</span><span class="sxs-lookup"><span data-stu-id="37ea2-740">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="37ea2-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="37ea2-742">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-742">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="37ea2-743">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="37ea2-743">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="37ea2-744">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-744">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-745">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-745">Parameters</span></span>
|<span data-ttu-id="37ea2-746">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-746">Name</span></span>|<span data-ttu-id="37ea2-747">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-747">Type</span></span>|<span data-ttu-id="37ea2-748">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-748">Attributes</span></span>|<span data-ttu-id="37ea2-749">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-749">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="37ea2-750">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-750">String</span></span>||<span data-ttu-id="37ea2-p136">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="37ea2-753">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-753">String</span></span>||<span data-ttu-id="37ea2-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="37ea2-756">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-756">Object</span></span>|<span data-ttu-id="37ea2-757">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-757">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-758">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-758">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-759">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-759">Object</span></span>|<span data-ttu-id="37ea2-760">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-760">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-761">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-761">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="37ea2-762">Booliano</span><span class="sxs-lookup"><span data-stu-id="37ea2-762">Boolean</span></span>|<span data-ttu-id="37ea2-763">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-763">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-764">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-764">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="37ea2-765">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-765">function</span></span>|<span data-ttu-id="37ea2-766">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-766">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-767">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-767">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37ea2-768">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-768">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="37ea2-769">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="37ea2-769">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37ea2-770">Erros</span><span class="sxs-lookup"><span data-stu-id="37ea2-770">Errors</span></span>

|<span data-ttu-id="37ea2-771">Código de erro</span><span class="sxs-lookup"><span data-stu-id="37ea2-771">Error code</span></span>|<span data-ttu-id="37ea2-772">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-772">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="37ea2-773">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="37ea2-773">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="37ea2-774">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="37ea2-774">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="37ea2-775">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-775">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-776">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-776">Requirements</span></span>

|<span data-ttu-id="37ea2-777">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-777">Requirement</span></span>|<span data-ttu-id="37ea2-778">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-778">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-779">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-779">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-780">1.1</span><span class="sxs-lookup"><span data-stu-id="37ea2-780">1.1</span></span>|
|[<span data-ttu-id="37ea2-781">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-781">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-782">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-782">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-783">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-783">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-784">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-784">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="37ea2-785">Exemplos</span><span class="sxs-lookup"><span data-stu-id="37ea2-785">Examples</span></span>

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

<span data-ttu-id="37ea2-786">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-786">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="37ea2-787">addFileAttachmentFromBase64Async (base64file, AttachmentName, [Options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-787">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="37ea2-788">Adiciona um arquivo da codificação Base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-788">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="37ea2-789">O `addFileAttachmentFromBase64Async` método carrega o arquivo da codificação Base64 e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="37ea2-789">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="37ea2-790">Esse método retorna o identificador de anexo no objeto AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="37ea2-790">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="37ea2-791">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-791">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-792">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-792">Parameters</span></span>

|<span data-ttu-id="37ea2-793">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-793">Name</span></span>|<span data-ttu-id="37ea2-794">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-794">Type</span></span>|<span data-ttu-id="37ea2-795">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-795">Attributes</span></span>|<span data-ttu-id="37ea2-796">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-796">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="37ea2-797">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-797">String</span></span>||<span data-ttu-id="37ea2-798">O conteúdo codificado em Base64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="37ea2-798">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="37ea2-799">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-799">String</span></span>||<span data-ttu-id="37ea2-p139">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="37ea2-802">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-802">Object</span></span>|<span data-ttu-id="37ea2-803">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-803">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-804">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-804">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-805">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-805">Object</span></span>|<span data-ttu-id="37ea2-806">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-806">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-807">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-807">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="37ea2-808">Booliano</span><span class="sxs-lookup"><span data-stu-id="37ea2-808">Boolean</span></span>|<span data-ttu-id="37ea2-809">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-809">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-810">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-810">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="37ea2-811">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-811">function</span></span>|<span data-ttu-id="37ea2-812">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-812">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-813">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-813">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37ea2-814">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-814">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="37ea2-815">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="37ea2-815">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37ea2-816">Erros</span><span class="sxs-lookup"><span data-stu-id="37ea2-816">Errors</span></span>

|<span data-ttu-id="37ea2-817">Código de erro</span><span class="sxs-lookup"><span data-stu-id="37ea2-817">Error code</span></span>|<span data-ttu-id="37ea2-818">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-818">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="37ea2-819">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="37ea2-819">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="37ea2-820">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="37ea2-820">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="37ea2-821">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-821">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-822">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-822">Requirements</span></span>

|<span data-ttu-id="37ea2-823">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-823">Requirement</span></span>|<span data-ttu-id="37ea2-824">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-825">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-826">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-826">Preview</span></span>|
|[<span data-ttu-id="37ea2-827">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-827">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-828">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-828">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-829">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-829">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-830">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-830">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="37ea2-831">Exemplos</span><span class="sxs-lookup"><span data-stu-id="37ea2-831">Examples</span></span>

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

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="37ea2-832">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-832">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="37ea2-833">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="37ea2-833">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="37ea2-834">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="37ea2-834">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-835">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-835">Parameters</span></span>

| <span data-ttu-id="37ea2-836">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-836">Name</span></span> | <span data-ttu-id="37ea2-837">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-837">Type</span></span> | <span data-ttu-id="37ea2-838">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-838">Attributes</span></span> | <span data-ttu-id="37ea2-839">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-839">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="37ea2-840">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="37ea2-840">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="37ea2-841">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="37ea2-841">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="37ea2-842">Função</span><span class="sxs-lookup"><span data-stu-id="37ea2-842">Function</span></span> || <span data-ttu-id="37ea2-p140">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="37ea2-846">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-846">Object</span></span> | <span data-ttu-id="37ea2-847">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-847">&lt;optional&gt;</span></span> | <span data-ttu-id="37ea2-848">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-848">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="37ea2-849">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-849">Object</span></span> | <span data-ttu-id="37ea2-850">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-850">&lt;optional&gt;</span></span> | <span data-ttu-id="37ea2-851">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-851">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="37ea2-852">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-852">function</span></span>| <span data-ttu-id="37ea2-853">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-853">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-854">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-854">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-855">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-855">Requirements</span></span>

|<span data-ttu-id="37ea2-856">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-856">Requirement</span></span>| <span data-ttu-id="37ea2-857">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-857">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-858">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-858">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37ea2-859">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-859">1.7</span></span> |
|[<span data-ttu-id="37ea2-860">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-860">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37ea2-861">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-861">ReadItem</span></span> |
|[<span data-ttu-id="37ea2-862">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-862">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37ea2-863">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-863">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="37ea2-864">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-864">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="37ea2-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="37ea2-866">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-866">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="37ea2-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="37ea2-870">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-870">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="37ea2-871">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-871">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-872">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-872">Parameters</span></span>

|<span data-ttu-id="37ea2-873">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-873">Name</span></span>|<span data-ttu-id="37ea2-874">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-874">Type</span></span>|<span data-ttu-id="37ea2-875">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-875">Attributes</span></span>|<span data-ttu-id="37ea2-876">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-876">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="37ea2-877">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-877">String</span></span>||<span data-ttu-id="37ea2-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="37ea2-880">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-880">String</span></span>||<span data-ttu-id="37ea2-881">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-881">The subject of the item to be attached.</span></span> <span data-ttu-id="37ea2-882">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-882">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="37ea2-883">Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-883">Object</span></span>|<span data-ttu-id="37ea2-884">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-884">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-885">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-885">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-886">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-886">Object</span></span>|<span data-ttu-id="37ea2-887">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-887">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-888">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-888">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-889">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-889">function</span></span>|<span data-ttu-id="37ea2-890">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-890">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-891">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-891">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37ea2-892">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-892">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="37ea2-893">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="37ea2-893">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37ea2-894">Erros</span><span class="sxs-lookup"><span data-stu-id="37ea2-894">Errors</span></span>

|<span data-ttu-id="37ea2-895">Código de erro</span><span class="sxs-lookup"><span data-stu-id="37ea2-895">Error code</span></span>|<span data-ttu-id="37ea2-896">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-896">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="37ea2-897">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-897">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-898">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-898">Requirements</span></span>

|<span data-ttu-id="37ea2-899">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-899">Requirement</span></span>|<span data-ttu-id="37ea2-900">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-900">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-901">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-901">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-902">1.1</span><span class="sxs-lookup"><span data-stu-id="37ea2-902">1.1</span></span>|
|[<span data-ttu-id="37ea2-903">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-903">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-904">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-904">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-905">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-905">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-906">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-906">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-907">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-907">Example</span></span>

<span data-ttu-id="37ea2-908">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-908">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="37ea2-909">close()</span><span class="sxs-lookup"><span data-stu-id="37ea2-909">close()</span></span>

<span data-ttu-id="37ea2-910">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-910">Closes the current item that is being composed.</span></span>

<span data-ttu-id="37ea2-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-913">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="37ea2-913">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="37ea2-914">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="37ea2-914">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-915">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-915">Requirements</span></span>

|<span data-ttu-id="37ea2-916">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-916">Requirement</span></span>|<span data-ttu-id="37ea2-917">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-917">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-918">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-918">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-919">1.3</span><span class="sxs-lookup"><span data-stu-id="37ea2-919">1.3</span></span>|
|[<span data-ttu-id="37ea2-920">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-920">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-921">Restrito</span><span class="sxs-lookup"><span data-stu-id="37ea2-921">Restricted</span></span>|
|[<span data-ttu-id="37ea2-922">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-922">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-923">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-923">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="37ea2-924">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-924">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="37ea2-925">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-925">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-926">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-926">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-927">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="37ea2-927">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="37ea2-928">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="37ea2-928">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="37ea2-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-932">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-932">Parameters</span></span>

|<span data-ttu-id="37ea2-933">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-933">Name</span></span>|<span data-ttu-id="37ea2-934">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-934">Type</span></span>|<span data-ttu-id="37ea2-935">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-935">Attributes</span></span>|<span data-ttu-id="37ea2-936">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-936">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="37ea2-937">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-937">String &#124; Object</span></span>||<span data-ttu-id="37ea2-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="37ea2-940">**OU**</span><span class="sxs-lookup"><span data-stu-id="37ea2-940">**OR**</span></span><br/><span data-ttu-id="37ea2-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="37ea2-943">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-943">String</span></span>|<span data-ttu-id="37ea2-944">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-944">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="37ea2-947">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-947">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="37ea2-948">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-948">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-949">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-949">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="37ea2-950">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-950">String</span></span>||<span data-ttu-id="37ea2-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="37ea2-953">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-953">String</span></span>||<span data-ttu-id="37ea2-954">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="37ea2-954">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="37ea2-955">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-955">String</span></span>||<span data-ttu-id="37ea2-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="37ea2-958">Booliano</span><span class="sxs-lookup"><span data-stu-id="37ea2-958">Boolean</span></span>||<span data-ttu-id="37ea2-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="37ea2-961">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-961">String</span></span>||<span data-ttu-id="37ea2-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="37ea2-965">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-965">function</span></span>|<span data-ttu-id="37ea2-966">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-966">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-967">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-968">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-968">Requirements</span></span>

|<span data-ttu-id="37ea2-969">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-969">Requirement</span></span>|<span data-ttu-id="37ea2-970">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-971">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-972">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-972">1.0</span></span>|
|[<span data-ttu-id="37ea2-973">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-974">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-975">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-976">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-976">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="37ea2-977">Exemplos</span><span class="sxs-lookup"><span data-stu-id="37ea2-977">Examples</span></span>

<span data-ttu-id="37ea2-978">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-978">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="37ea2-979">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="37ea2-979">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="37ea2-980">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-980">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="37ea2-981">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-981">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="37ea2-982">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-982">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="37ea2-983">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-983">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="37ea2-984">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-984">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="37ea2-985">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-985">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-986">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-986">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-987">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="37ea2-987">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="37ea2-988">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="37ea2-988">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="37ea2-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-992">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-992">Parameters</span></span>

|<span data-ttu-id="37ea2-993">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-993">Name</span></span>|<span data-ttu-id="37ea2-994">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-994">Type</span></span>|<span data-ttu-id="37ea2-995">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-995">Attributes</span></span>|<span data-ttu-id="37ea2-996">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-996">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="37ea2-997">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-997">String &#124; Object</span></span>||<span data-ttu-id="37ea2-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="37ea2-1000">**OU**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1000">**OR**</span></span><br/><span data-ttu-id="37ea2-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="37ea2-1003">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1003">String</span></span>|<span data-ttu-id="37ea2-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="37ea2-1007">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1007">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="37ea2-1008">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1009">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1009">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="37ea2-1010">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1010">String</span></span>||<span data-ttu-id="37ea2-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="37ea2-1013">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1013">String</span></span>||<span data-ttu-id="37ea2-1014">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1014">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="37ea2-1015">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1015">String</span></span>||<span data-ttu-id="37ea2-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="37ea2-1018">Booliano</span><span class="sxs-lookup"><span data-stu-id="37ea2-1018">Boolean</span></span>||<span data-ttu-id="37ea2-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="37ea2-1021">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1021">String</span></span>||<span data-ttu-id="37ea2-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1025">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1025">function</span></span>|<span data-ttu-id="37ea2-1026">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1027">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1028">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1028">Requirements</span></span>

|<span data-ttu-id="37ea2-1029">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1029">Requirement</span></span>|<span data-ttu-id="37ea2-1030">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1030">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1031">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1031">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1032">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1032">1.0</span></span>|
|[<span data-ttu-id="37ea2-1033">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1033">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1034">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1034">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1035">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1035">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1036">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1036">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="37ea2-1037">Exemplos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1037">Examples</span></span>

<span data-ttu-id="37ea2-1038">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1038">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="37ea2-1039">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1039">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="37ea2-1040">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1040">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="37ea2-1041">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1041">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="37ea2-1042">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1042">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="37ea2-1043">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1043">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="37ea2-1044">getAttachmentContentAsync (attachmentid, [opções], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1044">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="37ea2-1045">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um `AttachmentContent` objeto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1045">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="37ea2-1046">O `getAttachmentContentAsync` método obtém o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1046">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="37ea2-1047">Como prática recomendada, você deve usar o identificador para recuperar um anexo na mesma sessão em que o attachmentIds foi recuperado com a `getAttachmentsAsync` chamada ou `item.attachments` .</span><span class="sxs-lookup"><span data-stu-id="37ea2-1047">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="37ea2-1048">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1048">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="37ea2-1049">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1049">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1050">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1050">Parameters</span></span>

|<span data-ttu-id="37ea2-1051">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1051">Name</span></span>|<span data-ttu-id="37ea2-1052">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1052">Type</span></span>|<span data-ttu-id="37ea2-1053">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1053">Attributes</span></span>|<span data-ttu-id="37ea2-1054">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1054">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="37ea2-1055">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1055">String</span></span>||<span data-ttu-id="37ea2-1056">O identificador do anexo que você deseja obter.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1056">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="37ea2-1057">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1057">Object</span></span>|<span data-ttu-id="37ea2-1058">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1058">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1059">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1059">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1060">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1060">Object</span></span>|<span data-ttu-id="37ea2-1061">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1062">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1062">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1063">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1063">function</span></span>|<span data-ttu-id="37ea2-1064">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1065">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1065">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1066">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1066">Requirements</span></span>

|<span data-ttu-id="37ea2-1067">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1067">Requirement</span></span>|<span data-ttu-id="37ea2-1068">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1069">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1070">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-1070">Preview</span></span>|
|[<span data-ttu-id="37ea2-1071">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1072">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1073">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1074">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-1074">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1075">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1075">Returns:</span></span>

<span data-ttu-id="37ea2-1076">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1077">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1077">Example</span></span>

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

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="37ea2-1078">getAttachmentsAsync ([opções], [callback]) → array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37ea2-1078">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="37ea2-1079">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1079">Gets the item's attachments as an array.</span></span> <span data-ttu-id="37ea2-1080">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1080">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1081">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1081">Parameters</span></span>

|<span data-ttu-id="37ea2-1082">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1082">Name</span></span>|<span data-ttu-id="37ea2-1083">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1083">Type</span></span>|<span data-ttu-id="37ea2-1084">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1084">Attributes</span></span>|<span data-ttu-id="37ea2-1085">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1085">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="37ea2-1086">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1086">Object</span></span>|<span data-ttu-id="37ea2-1087">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1088">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1088">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1089">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1089">Object</span></span>|<span data-ttu-id="37ea2-1090">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1091">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1091">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1092">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1092">function</span></span>|<span data-ttu-id="37ea2-1093">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1093">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1094">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1094">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1095">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1095">Requirements</span></span>

|<span data-ttu-id="37ea2-1096">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1096">Requirement</span></span>|<span data-ttu-id="37ea2-1097">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1098">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1099">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-1099">Preview</span></span>|
|[<span data-ttu-id="37ea2-1100">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1101">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1102">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1103">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-1103">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1104">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1104">Returns:</span></span>

<span data-ttu-id="37ea2-1105">Tipo: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37ea2-1105">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1106">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1106">Example</span></span>

<span data-ttu-id="37ea2-1107">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1107">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="37ea2-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="37ea2-1109">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1109">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1110">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1110">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-1111">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1111">Requirements</span></span>

|<span data-ttu-id="37ea2-1112">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1112">Requirement</span></span>|<span data-ttu-id="37ea2-1113">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1114">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1115">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1115">1.0</span></span>|
|[<span data-ttu-id="37ea2-1116">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1116">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1117">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1117">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1118">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1118">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1119">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1119">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1120">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1120">Returns:</span></span>

<span data-ttu-id="37ea2-1121">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1121">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1122">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1122">Example</span></span>

<span data-ttu-id="37ea2-1123">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1123">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="37ea2-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="37ea2-1125">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1125">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1126">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1127">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1127">Parameters</span></span>

|<span data-ttu-id="37ea2-1128">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1128">Name</span></span>|<span data-ttu-id="37ea2-1129">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1129">Type</span></span>|<span data-ttu-id="37ea2-1130">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1130">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="37ea2-1131">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="37ea2-1131">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="37ea2-1132">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1132">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1133">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1133">Requirements</span></span>

|<span data-ttu-id="37ea2-1134">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1134">Requirement</span></span>|<span data-ttu-id="37ea2-1135">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1136">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1137">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1137">1.0</span></span>|
|[<span data-ttu-id="37ea2-1138">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1139">Restrito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1139">Restricted</span></span>|
|[<span data-ttu-id="37ea2-1140">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1141">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1142">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1142">Returns:</span></span>

<span data-ttu-id="37ea2-1143">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1143">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="37ea2-1144">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1144">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="37ea2-1145">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1145">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="37ea2-1146">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1146">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="37ea2-1147">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="37ea2-1147">Value of `entityType`</span></span>|<span data-ttu-id="37ea2-1148">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="37ea2-1148">Type of objects in returned array</span></span>|<span data-ttu-id="37ea2-1149">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="37ea2-1149">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="37ea2-1150">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1150">String</span></span>|<span data-ttu-id="37ea2-1151">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1151">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="37ea2-1152">Contato</span><span class="sxs-lookup"><span data-stu-id="37ea2-1152">Contact</span></span>|<span data-ttu-id="37ea2-1153">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1153">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="37ea2-1154">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1154">String</span></span>|<span data-ttu-id="37ea2-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1155">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="37ea2-1156">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="37ea2-1156">MeetingSuggestion</span></span>|<span data-ttu-id="37ea2-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1157">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="37ea2-1158">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="37ea2-1158">PhoneNumber</span></span>|<span data-ttu-id="37ea2-1159">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1159">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="37ea2-1160">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="37ea2-1160">TaskSuggestion</span></span>|<span data-ttu-id="37ea2-1161">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1161">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="37ea2-1162">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1162">String</span></span>|<span data-ttu-id="37ea2-1163">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="37ea2-1163">**Restricted**</span></span>|

<span data-ttu-id="37ea2-1164">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="37ea2-1164">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1165">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1165">Example</span></span>

<span data-ttu-id="37ea2-1166">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1166">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="37ea2-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="37ea2-1168">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1168">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1169">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1169">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-1170">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1170">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1171">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1171">Parameters</span></span>

|<span data-ttu-id="37ea2-1172">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1172">Name</span></span>|<span data-ttu-id="37ea2-1173">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1173">Type</span></span>|<span data-ttu-id="37ea2-1174">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1174">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="37ea2-1175">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1175">String</span></span>|<span data-ttu-id="37ea2-1176">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1176">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1177">Requirements</span></span>

|<span data-ttu-id="37ea2-1178">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1178">Requirement</span></span>|<span data-ttu-id="37ea2-1179">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1179">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1180">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1181">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1181">1.0</span></span>|
|[<span data-ttu-id="37ea2-1182">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1183">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1185">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1185">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1186">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1186">Returns:</span></span>

<span data-ttu-id="37ea2-p164">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="37ea2-1189">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="37ea2-1189">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="37ea2-1190">getInitializationContextAsync ([opções], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-1190">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="37ea2-1191">Obtém dados de inicialização passados quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1191">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1192">Este método só é compatível com o Outlook 2016 ou posterior no Windows (clique para executar versões posteriores a 16.0.8413.1000) e Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1192">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1193">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1193">Parameters</span></span>

|<span data-ttu-id="37ea2-1194">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1194">Name</span></span>|<span data-ttu-id="37ea2-1195">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1195">Type</span></span>|<span data-ttu-id="37ea2-1196">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1196">Attributes</span></span>|<span data-ttu-id="37ea2-1197">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1197">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="37ea2-1198">Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-1198">Object</span></span>|<span data-ttu-id="37ea2-1199">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1199">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1200">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1200">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1201">Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-1201">Object</span></span>|<span data-ttu-id="37ea2-1202">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1203">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1203">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1204">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1204">function</span></span>|<span data-ttu-id="37ea2-1205">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1206">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1206">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37ea2-1207">Com êxito, os dados de inicialização são fornecidos na `asyncResult.value` Propriedade como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1207">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="37ea2-1208">Se não houver nenhum contexto de inicialização, `asyncResult` o objeto conterá `Error` um objeto com `code` sua propriedade definida `9020` como e `name` sua propriedade definida `GenericResponseError`como.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1208">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1209">Requirements</span></span>

|<span data-ttu-id="37ea2-1210">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1210">Requirement</span></span>|<span data-ttu-id="37ea2-1211">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1213">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-1213">Preview</span></span>|
|[<span data-ttu-id="37ea2-1214">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1215">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1217">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-1218">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1218">Example</span></span>

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

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="37ea2-1219">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1219">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="37ea2-1220">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1220">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1221">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1221">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-p165">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="37ea2-1225">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1225">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="37ea2-1226">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1226">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="37ea2-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-1230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1230">Requirements</span></span>

|<span data-ttu-id="37ea2-1231">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1231">Requirement</span></span>|<span data-ttu-id="37ea2-1232">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1234">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1234">1.0</span></span>|
|[<span data-ttu-id="37ea2-1235">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1236">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1238">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1238">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1239">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1239">Returns:</span></span>

<span data-ttu-id="37ea2-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="37ea2-1242">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="37ea2-1242">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37ea2-1243">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1243">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37ea2-1244">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1244">Example</span></span>

<span data-ttu-id="37ea2-1245">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1245">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="37ea2-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="37ea2-1247">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1247">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1248">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1248">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-1249">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1249">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="37ea2-p168">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1252">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1252">Parameters</span></span>

|<span data-ttu-id="37ea2-1253">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1253">Name</span></span>|<span data-ttu-id="37ea2-1254">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1254">Type</span></span>|<span data-ttu-id="37ea2-1255">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1255">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="37ea2-1256">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="37ea2-1256">String</span></span>|<span data-ttu-id="37ea2-1257">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1257">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1258">Requirements</span></span>

|<span data-ttu-id="37ea2-1259">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1259">Requirement</span></span>|<span data-ttu-id="37ea2-1260">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1260">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1262">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1262">1.0</span></span>|
|[<span data-ttu-id="37ea2-1263">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1263">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1264">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1265">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1265">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1266">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1266">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1267">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1267">Returns:</span></span>

<span data-ttu-id="37ea2-1268">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1268">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="37ea2-1269">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="37ea2-1269">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37ea2-1270">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="37ea2-1270">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37ea2-1271">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1271">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="37ea2-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="37ea2-1273">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1273">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="37ea2-p169">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1276">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1276">Parameters</span></span>

|<span data-ttu-id="37ea2-1277">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1277">Name</span></span>|<span data-ttu-id="37ea2-1278">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1278">Type</span></span>|<span data-ttu-id="37ea2-1279">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1279">Attributes</span></span>|<span data-ttu-id="37ea2-1280">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1280">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="37ea2-1281">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="37ea2-1281">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="37ea2-p170">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="37ea2-1285">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1285">Object</span></span>|<span data-ttu-id="37ea2-1286">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1286">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1287">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1287">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1288">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1288">Object</span></span>|<span data-ttu-id="37ea2-1289">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1289">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1290">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1290">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1291">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1291">function</span></span>||<span data-ttu-id="37ea2-1292">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1292">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37ea2-1293">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1293">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="37ea2-1294">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1294">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1295">Requirements</span></span>

|<span data-ttu-id="37ea2-1296">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1296">Requirement</span></span>|<span data-ttu-id="37ea2-1297">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1299">1.2</span><span class="sxs-lookup"><span data-stu-id="37ea2-1299">1.2</span></span>|
|[<span data-ttu-id="37ea2-1300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1301">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1301">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-1302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1303">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-1303">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1304">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1304">Returns:</span></span>

<span data-ttu-id="37ea2-1305">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1305">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="37ea2-1306">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="37ea2-1306">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37ea2-1307">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1307">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37ea2-1308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1308">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="37ea2-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="37ea2-1310">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1310">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="37ea2-1311">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1311">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1312">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1312">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-1313">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1313">Requirements</span></span>

|<span data-ttu-id="37ea2-1314">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1314">Requirement</span></span>|<span data-ttu-id="37ea2-1315">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1315">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1316">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1317">1.6</span><span class="sxs-lookup"><span data-stu-id="37ea2-1317">1.6</span></span>|
|[<span data-ttu-id="37ea2-1318">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1319">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1320">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1321">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1321">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1322">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1322">Returns:</span></span>

<span data-ttu-id="37ea2-1323">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1323">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1324">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1324">Example</span></span>

<span data-ttu-id="37ea2-1325">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1325">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="37ea2-1326">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="37ea2-1326">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="37ea2-p173">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="37ea2-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1329">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1329">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37ea2-p174">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="37ea2-1333">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1333">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="37ea2-1334">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1334">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="37ea2-p175">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37ea2-1338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1338">Requirements</span></span>

|<span data-ttu-id="37ea2-1339">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1339">Requirement</span></span>|<span data-ttu-id="37ea2-1340">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1340">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1342">1.6</span><span class="sxs-lookup"><span data-stu-id="37ea2-1342">1.6</span></span>|
|[<span data-ttu-id="37ea2-1343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1344">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1345">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1346">Read</span><span class="sxs-lookup"><span data-stu-id="37ea2-1346">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37ea2-1347">Retorna:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1347">Returns:</span></span>

<span data-ttu-id="37ea2-p176">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="37ea2-1350">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1350">Example</span></span>

<span data-ttu-id="37ea2-1351">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1351">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="37ea2-1352">getSharedPropertiesAsync ([opções], retorno de chamada)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1352">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="37ea2-1353">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1353">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1354">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1354">Parameters</span></span>

|<span data-ttu-id="37ea2-1355">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1355">Name</span></span>|<span data-ttu-id="37ea2-1356">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1356">Type</span></span>|<span data-ttu-id="37ea2-1357">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1357">Attributes</span></span>|<span data-ttu-id="37ea2-1358">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1358">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="37ea2-1359">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1359">Object</span></span>|<span data-ttu-id="37ea2-1360">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1360">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1361">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1361">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1362">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1362">Object</span></span>|<span data-ttu-id="37ea2-1363">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1363">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1364">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1364">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1365">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1365">function</span></span>||<span data-ttu-id="37ea2-1366">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37ea2-1367">As propriedades compartilhadas são fornecidas [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) como um objeto `asyncResult.value` na propriedade.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1367">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="37ea2-1368">Este objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1368">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1369">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1369">Requirements</span></span>

|<span data-ttu-id="37ea2-1370">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1370">Requirement</span></span>|<span data-ttu-id="37ea2-1371">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1371">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1372">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1373">Visualização</span><span class="sxs-lookup"><span data-stu-id="37ea2-1373">Preview</span></span>|
|[<span data-ttu-id="37ea2-1374">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1375">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1376">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1377">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-1377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-1378">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1378">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="37ea2-1379">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="37ea2-1379">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="37ea2-1380">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1380">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="37ea2-p178">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1384">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1384">Parameters</span></span>

|<span data-ttu-id="37ea2-1385">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1385">Name</span></span>|<span data-ttu-id="37ea2-1386">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1386">Type</span></span>|<span data-ttu-id="37ea2-1387">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1387">Attributes</span></span>|<span data-ttu-id="37ea2-1388">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1388">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="37ea2-1389">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1389">function</span></span>||<span data-ttu-id="37ea2-1390">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1390">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37ea2-1391">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1391">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="37ea2-1392">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1392">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="37ea2-1393">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1393">Object</span></span>|<span data-ttu-id="37ea2-1394">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1394">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1395">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1395">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="37ea2-1396">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1396">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1397">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1397">Requirements</span></span>

|<span data-ttu-id="37ea2-1398">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1398">Requirement</span></span>|<span data-ttu-id="37ea2-1399">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1399">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1400">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1401">1.0</span><span class="sxs-lookup"><span data-stu-id="37ea2-1401">1.0</span></span>|
|[<span data-ttu-id="37ea2-1402">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1403">ReadItem</span></span>|
|[<span data-ttu-id="37ea2-1404">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1405">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-1405">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-1406">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1406">Example</span></span>

<span data-ttu-id="37ea2-p181">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="37ea2-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="37ea2-1411">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1411">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="37ea2-1412">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1412">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="37ea2-1413">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1413">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="37ea2-1414">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1414">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="37ea2-1415">Uma sessão termina quando o usuário fecha o aplicativo, ou se o usuário começa a redigir um formulário embutido e, em seguida, abre o formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1415">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1416">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1416">Parameters</span></span>

|<span data-ttu-id="37ea2-1417">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1417">Name</span></span>|<span data-ttu-id="37ea2-1418">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1418">Type</span></span>|<span data-ttu-id="37ea2-1419">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1419">Attributes</span></span>|<span data-ttu-id="37ea2-1420">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1420">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="37ea2-1421">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1421">String</span></span>||<span data-ttu-id="37ea2-1422">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1422">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="37ea2-1423">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1423">Object</span></span>|<span data-ttu-id="37ea2-1424">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1424">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1425">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1425">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1426">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1426">Object</span></span>|<span data-ttu-id="37ea2-1427">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1428">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1428">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1429">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1429">function</span></span>|<span data-ttu-id="37ea2-1430">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1431">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1431">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37ea2-1432">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1432">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37ea2-1433">Erros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1433">Errors</span></span>

|<span data-ttu-id="37ea2-1434">Código de erro</span><span class="sxs-lookup"><span data-stu-id="37ea2-1434">Error code</span></span>|<span data-ttu-id="37ea2-1435">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1435">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="37ea2-1436">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1436">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1437">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1437">Requirements</span></span>

|<span data-ttu-id="37ea2-1438">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1438">Requirement</span></span>|<span data-ttu-id="37ea2-1439">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1439">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1440">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1441">1.1</span><span class="sxs-lookup"><span data-stu-id="37ea2-1441">1.1</span></span>|
|[<span data-ttu-id="37ea2-1442">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1443">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1443">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-1444">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1445">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-1445">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-1446">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1446">Example</span></span>

<span data-ttu-id="37ea2-1447">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1447">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="37ea2-1448">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37ea2-1448">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="37ea2-1449">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1449">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="37ea2-1450">Atualmente, os tipos de eventos `Office.EventType.AttachmentsChanged`suportados `Office.EventType.AppointmentTimeChanged`são `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`,, `Office.EventType.RecurrenceChanged`e.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1450">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1451">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1451">Parameters</span></span>

| <span data-ttu-id="37ea2-1452">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1452">Name</span></span> | <span data-ttu-id="37ea2-1453">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1453">Type</span></span> | <span data-ttu-id="37ea2-1454">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1454">Attributes</span></span> | <span data-ttu-id="37ea2-1455">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1455">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="37ea2-1456">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="37ea2-1456">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="37ea2-1457">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1457">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="37ea2-1458">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1458">Object</span></span> | <span data-ttu-id="37ea2-1459">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1459">&lt;optional&gt;</span></span> | <span data-ttu-id="37ea2-1460">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1460">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="37ea2-1461">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1461">Object</span></span> | <span data-ttu-id="37ea2-1462">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1462">&lt;optional&gt;</span></span> | <span data-ttu-id="37ea2-1463">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1463">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="37ea2-1464">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1464">function</span></span>| <span data-ttu-id="37ea2-1465">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1466">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1467">Requirements</span></span>

|<span data-ttu-id="37ea2-1468">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1468">Requirement</span></span>| <span data-ttu-id="37ea2-1469">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37ea2-1471">1.7</span><span class="sxs-lookup"><span data-stu-id="37ea2-1471">1.7</span></span> |
|[<span data-ttu-id="37ea2-1472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37ea2-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1473">ReadItem</span></span> |
|[<span data-ttu-id="37ea2-1474">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="37ea2-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37ea2-1475">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="37ea2-1475">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="37ea2-1476">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1476">saveAsync([options], callback)</span></span>

<span data-ttu-id="37ea2-1477">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1477">Asynchronously saves an item.</span></span>

<span data-ttu-id="37ea2-p183">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1481">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1481">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="37ea2-1482">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1482">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="37ea2-p185">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="37ea2-1486">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="37ea2-1486">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="37ea2-1487">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1487">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="37ea2-1488">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1488">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="37ea2-1489">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1489">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1490">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1490">Parameters</span></span>

|<span data-ttu-id="37ea2-1491">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1491">Name</span></span>|<span data-ttu-id="37ea2-1492">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1492">Type</span></span>|<span data-ttu-id="37ea2-1493">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1493">Attributes</span></span>|<span data-ttu-id="37ea2-1494">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1494">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="37ea2-1495">Object</span><span class="sxs-lookup"><span data-stu-id="37ea2-1495">Object</span></span>|<span data-ttu-id="37ea2-1496">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1497">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1497">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1498">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1498">Object</span></span>|<span data-ttu-id="37ea2-1499">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1500">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1500">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1501">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1501">function</span></span>||<span data-ttu-id="37ea2-1502">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37ea2-1503">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1503">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1504">Requirements</span></span>

|<span data-ttu-id="37ea2-1505">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1505">Requirement</span></span>|<span data-ttu-id="37ea2-1506">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1506">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1507">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1508">1.3</span><span class="sxs-lookup"><span data-stu-id="37ea2-1508">1.3</span></span>|
|[<span data-ttu-id="37ea2-1509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1510">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1510">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-1511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1512">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-1512">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="37ea2-1513">Exemplos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1513">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="37ea2-p187">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="37ea2-1516">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="37ea2-1516">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="37ea2-1517">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1517">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="37ea2-p188">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37ea2-1521">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="37ea2-1521">Parameters</span></span>

|<span data-ttu-id="37ea2-1522">Nome</span><span class="sxs-lookup"><span data-stu-id="37ea2-1522">Name</span></span>|<span data-ttu-id="37ea2-1523">Tipo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1523">Type</span></span>|<span data-ttu-id="37ea2-1524">Atributos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1524">Attributes</span></span>|<span data-ttu-id="37ea2-1525">Descrição</span><span class="sxs-lookup"><span data-stu-id="37ea2-1525">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="37ea2-1526">String</span><span class="sxs-lookup"><span data-stu-id="37ea2-1526">String</span></span>||<span data-ttu-id="37ea2-p189">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="37ea2-1530">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1530">Object</span></span>|<span data-ttu-id="37ea2-1531">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1531">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1532">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="37ea2-1533">Objeto</span><span class="sxs-lookup"><span data-stu-id="37ea2-1533">Object</span></span>|<span data-ttu-id="37ea2-1534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-1535">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="37ea2-1536">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="37ea2-1536">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="37ea2-1537">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="37ea2-1537">&lt;optional&gt;</span></span>|<span data-ttu-id="37ea2-p190">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="37ea2-p191">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="37ea2-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="37ea2-1542">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="37ea2-1542">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="37ea2-1543">function</span><span class="sxs-lookup"><span data-stu-id="37ea2-1543">function</span></span>||<span data-ttu-id="37ea2-1544">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37ea2-1544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37ea2-1545">Requisitos</span><span class="sxs-lookup"><span data-stu-id="37ea2-1545">Requirements</span></span>

|<span data-ttu-id="37ea2-1546">Requisito</span><span class="sxs-lookup"><span data-stu-id="37ea2-1546">Requirement</span></span>|<span data-ttu-id="37ea2-1547">Valor</span><span class="sxs-lookup"><span data-stu-id="37ea2-1547">Value</span></span>|
|---|---|
|[<span data-ttu-id="37ea2-1548">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="37ea2-1548">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="37ea2-1549">1.2</span><span class="sxs-lookup"><span data-stu-id="37ea2-1549">1.2</span></span>|
|[<span data-ttu-id="37ea2-1550">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1550">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="37ea2-1551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37ea2-1551">ReadWriteItem</span></span>|
|[<span data-ttu-id="37ea2-1552">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="37ea2-1552">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="37ea2-1553">Escrever</span><span class="sxs-lookup"><span data-stu-id="37ea2-1553">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37ea2-1554">Exemplo</span><span class="sxs-lookup"><span data-stu-id="37ea2-1554">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
