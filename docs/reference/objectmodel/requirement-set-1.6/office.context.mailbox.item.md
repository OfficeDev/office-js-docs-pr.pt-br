---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,6
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: d6b77724290d9d100ff098baf11d97ba600bd8ee
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696033"
---
# <a name="item"></a><span data-ttu-id="d26f3-102">item</span><span class="sxs-lookup"><span data-stu-id="d26f3-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d26f3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d26f3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d26f3-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="d26f3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-106">Requirements</span></span>

|<span data-ttu-id="d26f3-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-107">Requirement</span></span>| <span data-ttu-id="d26f3-108">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-110">1.0</span></span>|
|[<span data-ttu-id="d26f3-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="d26f3-112">Restricted</span></span>|
|[<span data-ttu-id="d26f3-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d26f3-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="d26f3-115">Members and methods</span></span>

| <span data-ttu-id="d26f3-116">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-116">Member</span></span> | <span data-ttu-id="d26f3-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d26f3-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d26f3-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d26f3-119">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-119">Member</span></span> |
| [<span data-ttu-id="d26f3-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d26f3-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d26f3-121">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-121">Member</span></span> |
| [<span data-ttu-id="d26f3-122">body</span><span class="sxs-lookup"><span data-stu-id="d26f3-122">body</span></span>](#body-body) | <span data-ttu-id="d26f3-123">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-123">Member</span></span> |
| [<span data-ttu-id="d26f3-124">cc</span><span class="sxs-lookup"><span data-stu-id="d26f3-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d26f3-125">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-125">Member</span></span> |
| [<span data-ttu-id="d26f3-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="d26f3-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d26f3-127">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-127">Member</span></span> |
| [<span data-ttu-id="d26f3-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d26f3-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d26f3-129">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-129">Member</span></span> |
| [<span data-ttu-id="d26f3-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d26f3-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d26f3-131">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-131">Member</span></span> |
| [<span data-ttu-id="d26f3-132">end</span><span class="sxs-lookup"><span data-stu-id="d26f3-132">end</span></span>](#end-datetime) | <span data-ttu-id="d26f3-133">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-133">Member</span></span> |
| [<span data-ttu-id="d26f3-134">from</span><span class="sxs-lookup"><span data-stu-id="d26f3-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="d26f3-135">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-135">Member</span></span> |
| [<span data-ttu-id="d26f3-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d26f3-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d26f3-137">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-137">Member</span></span> |
| [<span data-ttu-id="d26f3-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="d26f3-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d26f3-139">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-139">Member</span></span> |
| [<span data-ttu-id="d26f3-140">itemId</span><span class="sxs-lookup"><span data-stu-id="d26f3-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d26f3-141">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-141">Member</span></span> |
| [<span data-ttu-id="d26f3-142">itemType</span><span class="sxs-lookup"><span data-stu-id="d26f3-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d26f3-143">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-143">Member</span></span> |
| [<span data-ttu-id="d26f3-144">location</span><span class="sxs-lookup"><span data-stu-id="d26f3-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="d26f3-145">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-145">Member</span></span> |
| [<span data-ttu-id="d26f3-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d26f3-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d26f3-147">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-147">Member</span></span> |
| [<span data-ttu-id="d26f3-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d26f3-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d26f3-149">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-149">Member</span></span> |
| [<span data-ttu-id="d26f3-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d26f3-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d26f3-151">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-151">Member</span></span> |
| [<span data-ttu-id="d26f3-152">organizer</span><span class="sxs-lookup"><span data-stu-id="d26f3-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="d26f3-153">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-153">Member</span></span> |
| [<span data-ttu-id="d26f3-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d26f3-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d26f3-155">Member</span><span class="sxs-lookup"><span data-stu-id="d26f3-155">Member</span></span> |
| [<span data-ttu-id="d26f3-156">sender</span><span class="sxs-lookup"><span data-stu-id="d26f3-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d26f3-157">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-157">Member</span></span> |
| [<span data-ttu-id="d26f3-158">start</span><span class="sxs-lookup"><span data-stu-id="d26f3-158">start</span></span>](#start-datetime) | <span data-ttu-id="d26f3-159">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-159">Member</span></span> |
| [<span data-ttu-id="d26f3-160">subject</span><span class="sxs-lookup"><span data-stu-id="d26f3-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d26f3-161">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-161">Member</span></span> |
| [<span data-ttu-id="d26f3-162">to</span><span class="sxs-lookup"><span data-stu-id="d26f3-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d26f3-163">Membro</span><span class="sxs-lookup"><span data-stu-id="d26f3-163">Member</span></span> |
| [<span data-ttu-id="d26f3-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d26f3-165">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-165">Method</span></span> |
| [<span data-ttu-id="d26f3-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d26f3-167">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-167">Method</span></span> |
| [<span data-ttu-id="d26f3-168">close</span><span class="sxs-lookup"><span data-stu-id="d26f3-168">close</span></span>](#close) | <span data-ttu-id="d26f3-169">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-169">Method</span></span> |
| [<span data-ttu-id="d26f3-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d26f3-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d26f3-171">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-171">Method</span></span> |
| [<span data-ttu-id="d26f3-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d26f3-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d26f3-173">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-173">Method</span></span> |
| [<span data-ttu-id="d26f3-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="d26f3-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d26f3-175">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-175">Method</span></span> |
| [<span data-ttu-id="d26f3-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d26f3-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d26f3-177">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-177">Method</span></span> |
| [<span data-ttu-id="d26f3-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d26f3-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d26f3-179">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-179">Method</span></span> |
| [<span data-ttu-id="d26f3-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d26f3-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d26f3-181">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-181">Method</span></span> |
| [<span data-ttu-id="d26f3-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d26f3-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d26f3-183">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-183">Method</span></span> |
| [<span data-ttu-id="d26f3-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d26f3-185">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-185">Method</span></span> |
| [<span data-ttu-id="d26f3-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="d26f3-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="d26f3-187">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-187">Method</span></span> |
| [<span data-ttu-id="d26f3-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d26f3-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d26f3-189">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-189">Method</span></span> |
| [<span data-ttu-id="d26f3-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d26f3-191">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-191">Method</span></span> |
| [<span data-ttu-id="d26f3-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d26f3-193">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-193">Method</span></span> |
| [<span data-ttu-id="d26f3-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d26f3-195">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-195">Method</span></span> |
| [<span data-ttu-id="d26f3-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d26f3-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d26f3-197">Método</span><span class="sxs-lookup"><span data-stu-id="d26f3-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d26f3-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-198">Example</span></span>

<span data-ttu-id="d26f3-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="d26f3-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d26f3-200">Membros</span><span class="sxs-lookup"><span data-stu-id="d26f3-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="d26f3-201">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="d26f3-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="d26f3-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="d26f3-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d26f3-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d26f3-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-206">Type</span></span>

*   <span data-ttu-id="d26f3-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="d26f3-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-208">Requirements</span></span>

|<span data-ttu-id="d26f3-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-209">Requirement</span></span>| <span data-ttu-id="d26f3-210">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-212">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-212">1.0</span></span>|
|[<span data-ttu-id="d26f3-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-214">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-216">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-217">Example</span></span>

<span data-ttu-id="d26f3-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="d26f3-219">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d26f3-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="d26f3-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-222">Type</span></span>

*   [<span data-ttu-id="d26f3-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="d26f3-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-224">Requirements</span></span>

|<span data-ttu-id="d26f3-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-225">Requirement</span></span>| <span data-ttu-id="d26f3-226">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-228">1.1</span><span class="sxs-lookup"><span data-stu-id="d26f3-228">1.1</span></span>|
|[<span data-ttu-id="d26f3-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-230">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="d26f3-234">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-236">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-236">Type</span></span>

*   [<span data-ttu-id="d26f3-237">Body</span><span class="sxs-lookup"><span data-stu-id="d26f3-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-238">Requirements</span></span>

|<span data-ttu-id="d26f3-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-239">Requirement</span></span>| <span data-ttu-id="d26f3-240">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-242">1.1</span><span class="sxs-lookup"><span data-stu-id="d26f3-242">1.1</span></span>|
|[<span data-ttu-id="d26f3-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-244">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-245">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-247">Example</span></span>

<span data-ttu-id="d26f3-248">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="d26f3-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d26f3-249">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="d26f3-250">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="d26f3-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-251">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d26f3-252">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-253">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-253">Read mode</span></span>

<span data-ttu-id="d26f3-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-256">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-256">Compose mode</span></span>

<span data-ttu-id="d26f3-257">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-258">Type</span></span>

*   <span data-ttu-id="d26f3-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-260">Requirements</span></span>

|<span data-ttu-id="d26f3-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-261">Requirement</span></span>| <span data-ttu-id="d26f3-262">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-264">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-264">1.0</span></span>|
|[<span data-ttu-id="d26f3-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-266">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d26f3-269">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="d26f3-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="d26f3-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="d26f3-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d26f3-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d26f3-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-275">Type</span></span>

*   <span data-ttu-id="d26f3-276">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-277">Requirements</span></span>

|<span data-ttu-id="d26f3-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-278">Requirement</span></span>| <span data-ttu-id="d26f3-279">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-281">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-281">1.0</span></span>|
|[<span data-ttu-id="d26f3-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-283">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d26f3-287">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="d26f3-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="d26f3-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-290">Type</span></span>

*   <span data-ttu-id="d26f3-291">Data</span><span class="sxs-lookup"><span data-stu-id="d26f3-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-292">Requirements</span></span>

|<span data-ttu-id="d26f3-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-293">Requirement</span></span>| <span data-ttu-id="d26f3-294">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-296">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-296">1.0</span></span>|
|[<span data-ttu-id="d26f3-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-298">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-300">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d26f3-302">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="d26f3-302">dateTimeModified: Date</span></span>

<span data-ttu-id="d26f3-303">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="d26f3-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="d26f3-304">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-305">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-306">Type</span></span>

*   <span data-ttu-id="d26f3-307">Data</span><span class="sxs-lookup"><span data-stu-id="d26f3-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-308">Requirements</span></span>

|<span data-ttu-id="d26f3-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-309">Requirement</span></span>| <span data-ttu-id="d26f3-310">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-312">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-312">1.0</span></span>|
|[<span data-ttu-id="d26f3-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-314">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-316">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="d26f3-318">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="d26f3-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d26f3-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-322">Read mode</span></span>

<span data-ttu-id="d26f3-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-324">Compose mode</span></span>

<span data-ttu-id="d26f3-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d26f3-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d26f3-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d26f3-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-328">Type</span></span>

*   <span data-ttu-id="d26f3-329">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-330">Requirements</span></span>

|<span data-ttu-id="d26f3-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-331">Requirement</span></span>| <span data-ttu-id="d26f3-332">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-334">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-334">1.0</span></span>|
|[<span data-ttu-id="d26f3-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-336">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="d26f3-339">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d26f3-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-345">Type</span></span>

*   [<span data-ttu-id="d26f3-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d26f3-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="d26f3-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="d26f3-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-348">Requirements</span></span>

|<span data-ttu-id="d26f3-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-349">Requirement</span></span>| <span data-ttu-id="d26f3-350">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-352">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-352">1.0</span></span>|
|[<span data-ttu-id="d26f3-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-354">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-356">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-356">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d26f3-357">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d26f3-357">internetMessageId: String</span></span>

<span data-ttu-id="d26f3-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-360">Type</span></span>

*   <span data-ttu-id="d26f3-361">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-362">Requirements</span></span>

|<span data-ttu-id="d26f3-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-363">Requirement</span></span>| <span data-ttu-id="d26f3-364">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-366">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-366">1.0</span></span>|
|[<span data-ttu-id="d26f3-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-368">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-370">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d26f3-372">doclass: String</span><span class="sxs-lookup"><span data-stu-id="d26f3-372">itemClass: String</span></span>

<span data-ttu-id="d26f3-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d26f3-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d26f3-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-377">Type</span></span> | <span data-ttu-id="d26f3-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-378">Description</span></span> | <span data-ttu-id="d26f3-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="d26f3-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d26f3-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="d26f3-380">Appointment items</span></span> | <span data-ttu-id="d26f3-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="d26f3-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="d26f3-382">Message items</span></span> | <span data-ttu-id="d26f3-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="d26f3-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d26f3-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-385">Type</span></span>

*   <span data-ttu-id="d26f3-386">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-387">Requirements</span></span>

|<span data-ttu-id="d26f3-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-388">Requirement</span></span>| <span data-ttu-id="d26f3-389">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-391">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-391">1.0</span></span>|
|[<span data-ttu-id="d26f3-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-393">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-395">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d26f3-397">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d26f3-397">(nullable) itemId: String</span></span>

<span data-ttu-id="d26f3-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-400">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="d26f3-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d26f3-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d26f3-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d26f3-402">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="d26f3-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d26f3-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d26f3-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d26f3-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-406">Type</span></span>

*   <span data-ttu-id="d26f3-407">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-408">Requirements</span></span>

|<span data-ttu-id="d26f3-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-409">Requirement</span></span>| <span data-ttu-id="d26f3-410">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-412">1.0</span></span>|
|[<span data-ttu-id="d26f3-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-414">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-416">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-417">Example</span></span>

<span data-ttu-id="d26f3-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="d26f3-420">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-421">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="d26f3-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d26f3-422">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-423">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-423">Type</span></span>

*   [<span data-ttu-id="d26f3-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d26f3-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-425">Requirements</span></span>

|<span data-ttu-id="d26f3-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-426">Requirement</span></span>| <span data-ttu-id="d26f3-427">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-429">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-429">1.0</span></span>|
|[<span data-ttu-id="d26f3-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-431">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-432">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-433">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-434">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="d26f3-435">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-436">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-437">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-437">Read mode</span></span>

<span data-ttu-id="d26f3-438">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-439">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-439">Compose mode</span></span>

<span data-ttu-id="d26f3-440">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-441">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-441">Type</span></span>

*   <span data-ttu-id="d26f3-442">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-443">Requirements</span></span>

|<span data-ttu-id="d26f3-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-444">Requirement</span></span>| <span data-ttu-id="d26f3-445">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-447">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-447">1.0</span></span>|
|[<span data-ttu-id="d26f3-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-449">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-450">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-451">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-451">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d26f3-452">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d26f3-452">normalizedSubject: String</span></span>

<span data-ttu-id="d26f3-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d26f3-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="d26f3-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-457">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-457">Type</span></span>

*   <span data-ttu-id="d26f3-458">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-459">Requirements</span></span>

|<span data-ttu-id="d26f3-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-460">Requirement</span></span>| <span data-ttu-id="d26f3-461">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-463">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-463">1.0</span></span>|
|[<span data-ttu-id="d26f3-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-465">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-467">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-468">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="d26f3-469">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-470">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-471">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-471">Type</span></span>

*   [<span data-ttu-id="d26f3-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d26f3-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-473">Requirements</span></span>

|<span data-ttu-id="d26f3-474">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-474">Requirement</span></span>| <span data-ttu-id="d26f3-475">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-477">1.3</span><span class="sxs-lookup"><span data-stu-id="d26f3-477">1.3</span></span>|
|[<span data-ttu-id="d26f3-478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-479">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-480">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-481">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-482">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="d26f3-483">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="d26f3-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-484">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="d26f3-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d26f3-485">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-486">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-486">Read mode</span></span>

<span data-ttu-id="d26f3-487">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="d26f3-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-488">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-488">Compose mode</span></span>

<span data-ttu-id="d26f3-489">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="d26f3-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-490">Type</span></span>

*   <span data-ttu-id="d26f3-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-492">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-492">Requirements</span></span>

|<span data-ttu-id="d26f3-493">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-493">Requirement</span></span>| <span data-ttu-id="d26f3-494">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-495">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-496">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-496">1.0</span></span>|
|[<span data-ttu-id="d26f3-497">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-498">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-499">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-500">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-500">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="d26f3-501">organizador: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-504">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-504">Type</span></span>

*   [<span data-ttu-id="d26f3-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d26f3-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-506">Requirements</span></span>

|<span data-ttu-id="d26f3-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-507">Requirement</span></span>| <span data-ttu-id="d26f3-508">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-510">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-510">1.0</span></span>|
|[<span data-ttu-id="d26f3-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-512">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-514">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-515">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="d26f3-516">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="d26f3-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-517">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="d26f3-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d26f3-518">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-519">Read mode</span></span>

<span data-ttu-id="d26f3-520">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="d26f3-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-521">Compose mode</span></span>

<span data-ttu-id="d26f3-522">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="d26f3-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-523">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-523">Type</span></span>

*   <span data-ttu-id="d26f3-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-525">Requirements</span></span>

|<span data-ttu-id="d26f3-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-526">Requirement</span></span>| <span data-ttu-id="d26f3-527">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-528">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-529">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-529">1.0</span></span>|
|[<span data-ttu-id="d26f3-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-531">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-532">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-533">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-533">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="d26f3-534">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d26f3-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-539">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d26f3-540">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-540">Type</span></span>

*   [<span data-ttu-id="d26f3-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d26f3-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="d26f3-542">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-542">Requirements</span></span>

|<span data-ttu-id="d26f3-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-543">Requirement</span></span>| <span data-ttu-id="d26f3-544">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-546">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-546">1.0</span></span>|
|[<span data-ttu-id="d26f3-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-548">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-550">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-551">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-551">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="d26f3-552">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-553">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="d26f3-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d26f3-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-556">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-556">Read mode</span></span>

<span data-ttu-id="d26f3-557">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-557">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-558">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-558">Compose mode</span></span>

<span data-ttu-id="d26f3-559">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d26f3-560">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d26f3-561">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d26f3-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-562">Type</span></span>

*   <span data-ttu-id="d26f3-563">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-564">Requirements</span></span>

|<span data-ttu-id="d26f3-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-565">Requirement</span></span>| <span data-ttu-id="d26f3-566">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-568">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-568">1.0</span></span>|
|[<span data-ttu-id="d26f3-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-570">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="d26f3-573">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-574">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d26f3-575">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="d26f3-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-576">Read mode</span></span>

<span data-ttu-id="d26f3-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-579">Compose mode</span></span>

<span data-ttu-id="d26f3-580">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="d26f3-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-581">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-581">Type</span></span>

*   <span data-ttu-id="d26f3-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-583">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-583">Requirements</span></span>

|<span data-ttu-id="d26f3-584">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-584">Requirement</span></span>| <span data-ttu-id="d26f3-585">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-586">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-587">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-587">1.0</span></span>|
|[<span data-ttu-id="d26f3-588">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-589">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-590">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-591">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-591">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="d26f3-592">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d26f3-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="d26f3-593">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d26f3-594">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d26f3-595">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="d26f3-595">Read mode</span></span>

<span data-ttu-id="d26f3-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d26f3-598">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="d26f3-598">Compose mode</span></span>

<span data-ttu-id="d26f3-599">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d26f3-600">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-600">Type</span></span>

*   <span data-ttu-id="d26f3-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-602">Requirements</span></span>

|<span data-ttu-id="d26f3-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-603">Requirement</span></span>| <span data-ttu-id="d26f3-604">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-606">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-606">1.0</span></span>|
|[<span data-ttu-id="d26f3-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-608">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-609">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-610">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d26f3-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="d26f3-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d26f3-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d26f3-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d26f3-613">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d26f3-614">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="d26f3-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d26f3-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="d26f3-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-616">Parameters</span></span>

|<span data-ttu-id="d26f3-617">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-617">Name</span></span>| <span data-ttu-id="d26f3-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-618">Type</span></span>| <span data-ttu-id="d26f3-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-619">Attributes</span></span>| <span data-ttu-id="d26f3-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d26f3-621">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-621">String</span></span>||<span data-ttu-id="d26f3-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d26f3-624">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-624">String</span></span>||<span data-ttu-id="d26f3-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d26f3-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-627">Object</span></span>| <span data-ttu-id="d26f3-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-628">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-629">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="d26f3-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-630">Object</span></span> | <span data-ttu-id="d26f3-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-631">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="d26f3-633">Booliano</span><span class="sxs-lookup"><span data-stu-id="d26f3-633">Boolean</span></span> | <span data-ttu-id="d26f3-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-634">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-635">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="d26f3-636">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-636">function</span></span>| <span data-ttu-id="d26f3-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-637">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d26f3-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d26f3-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="d26f3-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d26f3-641">Erros</span><span class="sxs-lookup"><span data-stu-id="d26f3-641">Errors</span></span>

| <span data-ttu-id="d26f3-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d26f3-642">Error code</span></span> | <span data-ttu-id="d26f3-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d26f3-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="d26f3-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d26f3-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="d26f3-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d26f3-646">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-647">Requirements</span></span>

|<span data-ttu-id="d26f3-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-648">Requirement</span></span>| <span data-ttu-id="d26f3-649">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-651">1.1</span><span class="sxs-lookup"><span data-stu-id="d26f3-651">1.1</span></span>|
|[<span data-ttu-id="d26f3-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-655">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d26f3-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="d26f3-656">Examples</span></span>

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

<span data-ttu-id="d26f3-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d26f3-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d26f3-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d26f3-659">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d26f3-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d26f3-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="d26f3-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d26f3-664">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-665">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-665">Parameters</span></span>

|<span data-ttu-id="d26f3-666">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-666">Name</span></span>| <span data-ttu-id="d26f3-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-667">Type</span></span>| <span data-ttu-id="d26f3-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-668">Attributes</span></span>| <span data-ttu-id="d26f3-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d26f3-670">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-670">String</span></span>||<span data-ttu-id="d26f3-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d26f3-673">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d26f3-673">String</span></span>||<span data-ttu-id="d26f3-674">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-674">The subject of the item to be attached.</span></span> <span data-ttu-id="d26f3-675">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d26f3-676">Object</span><span class="sxs-lookup"><span data-stu-id="d26f3-676">Object</span></span>| <span data-ttu-id="d26f3-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-677">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d26f3-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-679">Object</span></span>| <span data-ttu-id="d26f3-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-680">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d26f3-682">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-682">function</span></span>| <span data-ttu-id="d26f3-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-683">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d26f3-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d26f3-686">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="d26f3-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d26f3-687">Erros</span><span class="sxs-lookup"><span data-stu-id="d26f3-687">Errors</span></span>

| <span data-ttu-id="d26f3-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d26f3-688">Error code</span></span> | <span data-ttu-id="d26f3-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d26f3-690">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-691">Requirements</span></span>

|<span data-ttu-id="d26f3-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-692">Requirement</span></span>| <span data-ttu-id="d26f3-693">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-695">1.1</span><span class="sxs-lookup"><span data-stu-id="d26f3-695">1.1</span></span>|
|[<span data-ttu-id="d26f3-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-699">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-700">Example</span></span>

<span data-ttu-id="d26f3-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d26f3-702">close()</span><span class="sxs-lookup"><span data-stu-id="d26f3-702">close()</span></span>

<span data-ttu-id="d26f3-703">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="d26f3-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d26f3-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-706">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="d26f3-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d26f3-707">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="d26f3-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-708">Requirements</span></span>

|<span data-ttu-id="d26f3-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-709">Requirement</span></span>| <span data-ttu-id="d26f3-710">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-712">1.3</span><span class="sxs-lookup"><span data-stu-id="d26f3-712">1.3</span></span>|
|[<span data-ttu-id="d26f3-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="d26f3-714">Restricted</span></span>|
|[<span data-ttu-id="d26f3-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-716">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-716">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d26f3-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d26f3-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d26f3-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-719">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-720">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="d26f3-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d26f3-721">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d26f3-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d26f3-722">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="d26f3-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="d26f3-723">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="d26f3-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="d26f3-724">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-725">Parameters</span></span>

| <span data-ttu-id="d26f3-726">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-726">Name</span></span> | <span data-ttu-id="d26f3-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-727">Type</span></span> | <span data-ttu-id="d26f3-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-728">Attributes</span></span> | <span data-ttu-id="d26f3-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="d26f3-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d26f3-730">String &#124; Object</span></span>| |<span data-ttu-id="d26f3-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d26f3-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="d26f3-733">**OR**</span></span><br/><span data-ttu-id="d26f3-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d26f3-736">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-736">String</span></span> | <span data-ttu-id="d26f3-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-737">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d26f3-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d26f3-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-741">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-742">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d26f3-743">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-743">String</span></span> | | <span data-ttu-id="d26f3-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d26f3-746">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d26f3-746">String</span></span> | | <span data-ttu-id="d26f3-747">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="d26f3-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d26f3-748">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-748">String</span></span> | | <span data-ttu-id="d26f3-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="d26f3-751">Booliano</span><span class="sxs-lookup"><span data-stu-id="d26f3-751">Boolean</span></span> | | <span data-ttu-id="d26f3-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d26f3-754">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-754">String</span></span> | | <span data-ttu-id="d26f3-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d26f3-758">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-758">function</span></span> | <span data-ttu-id="d26f3-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-759">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-761">Requirements</span></span>

|<span data-ttu-id="d26f3-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-762">Requirement</span></span>| <span data-ttu-id="d26f3-763">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-765">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-765">1.0</span></span>|
|[<span data-ttu-id="d26f3-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-767">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-769">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d26f3-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="d26f3-770">Examples</span></span>

<span data-ttu-id="d26f3-771">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d26f3-772">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="d26f3-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d26f3-773">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d26f3-774">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d26f3-775">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d26f3-776">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d26f3-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d26f3-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d26f3-778">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-779">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-780">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="d26f3-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d26f3-781">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d26f3-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d26f3-782">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="d26f3-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="d26f3-783">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="d26f3-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="d26f3-784">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-785">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-785">Parameters</span></span>

| <span data-ttu-id="d26f3-786">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-786">Name</span></span> | <span data-ttu-id="d26f3-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-787">Type</span></span> | <span data-ttu-id="d26f3-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-788">Attributes</span></span> | <span data-ttu-id="d26f3-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="d26f3-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d26f3-790">String &#124; Object</span></span>| | <span data-ttu-id="d26f3-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d26f3-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="d26f3-793">**OR**</span></span><br/><span data-ttu-id="d26f3-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d26f3-796">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-796">String</span></span> | <span data-ttu-id="d26f3-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-797">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d26f3-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d26f3-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-801">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-802">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d26f3-803">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-803">String</span></span> | | <span data-ttu-id="d26f3-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d26f3-806">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d26f3-806">String</span></span> | | <span data-ttu-id="d26f3-807">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="d26f3-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d26f3-808">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-808">String</span></span> | | <span data-ttu-id="d26f3-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="d26f3-811">Booliano</span><span class="sxs-lookup"><span data-stu-id="d26f3-811">Boolean</span></span> | | <span data-ttu-id="d26f3-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d26f3-814">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-814">String</span></span> | | <span data-ttu-id="d26f3-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d26f3-818">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-818">function</span></span> | <span data-ttu-id="d26f3-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-819">&lt;optional&gt;</span></span> | <span data-ttu-id="d26f3-820">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-821">Requirements</span></span>

|<span data-ttu-id="d26f3-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-822">Requirement</span></span>| <span data-ttu-id="d26f3-823">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-824">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-825">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-825">1.0</span></span>|
|[<span data-ttu-id="d26f3-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-827">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-829">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d26f3-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="d26f3-830">Examples</span></span>

<span data-ttu-id="d26f3-831">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d26f3-832">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="d26f3-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d26f3-833">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d26f3-834">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d26f3-835">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d26f3-836">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="d26f3-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="d26f3-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="d26f3-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-839">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-840">Requirements</span></span>

|<span data-ttu-id="d26f3-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-841">Requirement</span></span>| <span data-ttu-id="d26f3-842">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-844">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-844">1.0</span></span>|
|[<span data-ttu-id="d26f3-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-846">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-848">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-849">Returns:</span></span>

<span data-ttu-id="d26f3-850">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-851">Example</span></span>

<span data-ttu-id="d26f3-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-852">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="d26f3-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="d26f3-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="d26f3-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-855">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-856">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-856">Parameters</span></span>

|<span data-ttu-id="d26f3-857">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-857">Name</span></span>| <span data-ttu-id="d26f3-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-858">Type</span></span>| <span data-ttu-id="d26f3-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d26f3-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d26f3-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="d26f3-861">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="d26f3-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-862">Requirements</span></span>

|<span data-ttu-id="d26f3-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-863">Requirement</span></span>| <span data-ttu-id="d26f3-864">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-865">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-866">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-866">1.0</span></span>|
|[<span data-ttu-id="d26f3-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="d26f3-868">Restricted</span></span>|
|[<span data-ttu-id="d26f3-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-870">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-871">Returns:</span></span>

<span data-ttu-id="d26f3-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d26f3-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="d26f3-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d26f3-874">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d26f3-875">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d26f3-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="d26f3-876">Value of `entityType`</span></span> | <span data-ttu-id="d26f3-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="d26f3-877">Type of objects in returned array</span></span> | <span data-ttu-id="d26f3-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="d26f3-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d26f3-879">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-879">String</span></span> | <span data-ttu-id="d26f3-880">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="d26f3-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d26f3-881">Contato</span><span class="sxs-lookup"><span data-stu-id="d26f3-881">Contact</span></span> | <span data-ttu-id="d26f3-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d26f3-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d26f3-883">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-883">String</span></span> | <span data-ttu-id="d26f3-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d26f3-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d26f3-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d26f3-885">MeetingSuggestion</span></span> | <span data-ttu-id="d26f3-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d26f3-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d26f3-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d26f3-887">PhoneNumber</span></span> | <span data-ttu-id="d26f3-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="d26f3-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d26f3-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d26f3-889">TaskSuggestion</span></span> | <span data-ttu-id="d26f3-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d26f3-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d26f3-891">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-891">String</span></span> | <span data-ttu-id="d26f3-892">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="d26f3-892">**Restricted**</span></span> |

<span data-ttu-id="d26f3-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="d26f3-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-894">Example</span></span>

<span data-ttu-id="d26f3-895">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="d26f3-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="d26f3-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="d26f3-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="d26f3-897">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-898">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-900">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-900">Parameters</span></span>

|<span data-ttu-id="d26f3-901">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-901">Name</span></span>| <span data-ttu-id="d26f3-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-902">Type</span></span>| <span data-ttu-id="d26f3-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d26f3-904">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-904">String</span></span>|<span data-ttu-id="d26f3-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="d26f3-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-906">Requirements</span></span>

|<span data-ttu-id="d26f3-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-907">Requirement</span></span>| <span data-ttu-id="d26f3-908">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-909">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-910">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-910">1.0</span></span>|
|[<span data-ttu-id="d26f3-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-912">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-913">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-914">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-915">Returns:</span></span>

<span data-ttu-id="d26f3-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d26f3-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="d26f3-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d26f3-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d26f3-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d26f3-920">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-921">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d26f3-925">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="d26f3-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d26f3-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d26f3-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-930">Requirements</span></span>

|<span data-ttu-id="d26f3-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-931">Requirement</span></span>| <span data-ttu-id="d26f3-932">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-934">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-934">1.0</span></span>|
|[<span data-ttu-id="d26f3-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-936">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-938">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-939">Returns:</span></span>

<span data-ttu-id="d26f3-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="d26f3-942">Tipo: objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-942">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-943">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-943">Example</span></span>

<span data-ttu-id="d26f3-944">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="d26f3-944">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d26f3-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d26f3-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d26f3-946">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-946">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-947">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-948">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-948">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d26f3-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-951">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-951">Parameters</span></span>

|<span data-ttu-id="d26f3-952">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-952">Name</span></span>| <span data-ttu-id="d26f3-953">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-953">Type</span></span>| <span data-ttu-id="d26f3-954">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-954">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d26f3-955">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-955">String</span></span>|<span data-ttu-id="d26f3-956">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="d26f3-956">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-957">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-957">Requirements</span></span>

|<span data-ttu-id="d26f3-958">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-958">Requirement</span></span>| <span data-ttu-id="d26f3-959">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-959">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-960">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-960">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-961">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-961">1.0</span></span>|
|[<span data-ttu-id="d26f3-962">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-962">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-963">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-963">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-964">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-964">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-965">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-965">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-966">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-966">Returns:</span></span>

<span data-ttu-id="d26f3-967">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-967">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d26f3-968">Tipo: cadeia de caracteres de matriz. < ></span><span class="sxs-lookup"><span data-stu-id="d26f3-968">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-969">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-969">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d26f3-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d26f3-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d26f3-971">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-971">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d26f3-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-974">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-974">Parameters</span></span>

|<span data-ttu-id="d26f3-975">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-975">Name</span></span>| <span data-ttu-id="d26f3-976">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-976">Type</span></span>| <span data-ttu-id="d26f3-977">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-977">Attributes</span></span>| <span data-ttu-id="d26f3-978">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-978">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d26f3-979">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d26f3-979">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d26f3-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d26f3-983">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-983">Object</span></span>| <span data-ttu-id="d26f3-984">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-984">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-985">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-985">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d26f3-986">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-986">Object</span></span>| <span data-ttu-id="d26f3-987">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-987">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-988">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-988">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d26f3-989">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-989">function</span></span>||<span data-ttu-id="d26f3-990">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d26f3-991">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-991">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d26f3-992">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-992">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-993">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-993">Requirements</span></span>

|<span data-ttu-id="d26f3-994">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-994">Requirement</span></span>| <span data-ttu-id="d26f3-995">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-995">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-996">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-996">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-997">1.2</span><span class="sxs-lookup"><span data-stu-id="d26f3-997">1.2</span></span>|
|[<span data-ttu-id="d26f3-998">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-998">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-999">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-999">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-1000">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-1000">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1001">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-1001">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-1002">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-1002">Returns:</span></span>

<span data-ttu-id="d26f3-1003">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1003">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d26f3-1004">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="d26f3-1004">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-1005">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1005">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="d26f3-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="d26f3-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="d26f3-1007">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1007">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="d26f3-1008">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d26f3-1008">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-1009">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1009">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-1010">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1010">Requirements</span></span>

|<span data-ttu-id="d26f3-1011">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1011">Requirement</span></span>| <span data-ttu-id="d26f3-1012">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1013">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1014">1.6</span><span class="sxs-lookup"><span data-stu-id="d26f3-1014">1.6</span></span> |
|[<span data-ttu-id="d26f3-1015">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1016">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-1017">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1018">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-1018">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-1019">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-1019">Returns:</span></span>

<span data-ttu-id="d26f3-1020">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="d26f3-1020">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-1021">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1021">Example</span></span>

<span data-ttu-id="d26f3-1022">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1022">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d26f3-1023">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d26f3-1023">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d26f3-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d26f3-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-1026">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1026">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d26f3-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d26f3-1030">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="d26f3-1030">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d26f3-1031">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1031">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d26f3-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d26f3-1035">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1035">Requirements</span></span>

|<span data-ttu-id="d26f3-1036">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1036">Requirement</span></span>| <span data-ttu-id="d26f3-1037">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1037">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1038">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1038">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1039">1.6</span><span class="sxs-lookup"><span data-stu-id="d26f3-1039">1.6</span></span> |
|[<span data-ttu-id="d26f3-1040">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1040">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1041">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1041">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-1042">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-1042">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1043">Read</span><span class="sxs-lookup"><span data-stu-id="d26f3-1043">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d26f3-1044">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d26f3-1044">Returns:</span></span>

<span data-ttu-id="d26f3-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d26f3-1047">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1047">Example</span></span>

<span data-ttu-id="d26f3-1048">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1048">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d26f3-1049">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d26f3-1049">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d26f3-1050">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1050">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d26f3-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-1054">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-1054">Parameters</span></span>

|<span data-ttu-id="d26f3-1055">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-1055">Name</span></span>| <span data-ttu-id="d26f3-1056">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1056">Type</span></span>| <span data-ttu-id="d26f3-1057">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1057">Attributes</span></span>| <span data-ttu-id="d26f3-1058">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-1058">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d26f3-1059">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-1059">function</span></span>||<span data-ttu-id="d26f3-1060">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-1060">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d26f3-1061">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1061">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d26f3-1062">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1062">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d26f3-1063">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1063">Object</span></span>| <span data-ttu-id="d26f3-1064">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1065">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1065">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d26f3-1066">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1066">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-1067">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1067">Requirements</span></span>

|<span data-ttu-id="d26f3-1068">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1068">Requirement</span></span>| <span data-ttu-id="d26f3-1069">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1069">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1070">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1070">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1071">1.0</span><span class="sxs-lookup"><span data-stu-id="d26f3-1071">1.0</span></span>|
|[<span data-ttu-id="d26f3-1072">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1072">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1073">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1073">ReadItem</span></span>|
|[<span data-ttu-id="d26f3-1074">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d26f3-1074">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1075">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d26f3-1075">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-1076">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1076">Example</span></span>

<span data-ttu-id="d26f3-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d26f3-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d26f3-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d26f3-1081">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1081">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d26f3-1082">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1082">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d26f3-1083">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1083">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d26f3-1084">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1084">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d26f3-1085">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1085">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-1086">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-1086">Parameters</span></span>

|<span data-ttu-id="d26f3-1087">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-1087">Name</span></span>| <span data-ttu-id="d26f3-1088">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1088">Type</span></span>| <span data-ttu-id="d26f3-1089">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1089">Attributes</span></span>| <span data-ttu-id="d26f3-1090">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-1090">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d26f3-1091">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-1091">String</span></span>||<span data-ttu-id="d26f3-1092">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1092">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="d26f3-1093">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1093">Object</span></span>| <span data-ttu-id="d26f3-1094">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1095">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d26f3-1096">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1096">Object</span></span>| <span data-ttu-id="d26f3-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1098">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d26f3-1099">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-1099">function</span></span>| <span data-ttu-id="d26f3-1100">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1101">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d26f3-1102">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1102">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d26f3-1103">Erros</span><span class="sxs-lookup"><span data-stu-id="d26f3-1103">Errors</span></span>

| <span data-ttu-id="d26f3-1104">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d26f3-1104">Error code</span></span> | <span data-ttu-id="d26f3-1105">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-1105">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d26f3-1106">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1106">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-1107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1107">Requirements</span></span>

|<span data-ttu-id="d26f3-1108">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1108">Requirement</span></span>| <span data-ttu-id="d26f3-1109">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1111">1.1</span><span class="sxs-lookup"><span data-stu-id="d26f3-1111">1.1</span></span>|
|[<span data-ttu-id="d26f3-1112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1113">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1113">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-1114">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1115">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-1115">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-1116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1116">Example</span></span>

<span data-ttu-id="d26f3-1117">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1117">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d26f3-1118">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d26f3-1118">saveAsync([options], callback)</span></span>

<span data-ttu-id="d26f3-1119">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1119">Asynchronously saves an item.</span></span>

<span data-ttu-id="d26f3-1120">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1120">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="d26f3-1121">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1121">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="d26f3-1122">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1122">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-1123">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1123">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d26f3-1124">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1124">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d26f3-p175">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d26f3-1128">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="d26f3-1128">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d26f3-1129">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1129">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="d26f3-1130">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1130">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="d26f3-1131">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1131">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d26f3-1132">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1132">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-1133">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-1133">Parameters</span></span>

|<span data-ttu-id="d26f3-1134">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-1134">Name</span></span>| <span data-ttu-id="d26f3-1135">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1135">Type</span></span>| <span data-ttu-id="d26f3-1136">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1136">Attributes</span></span>| <span data-ttu-id="d26f3-1137">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-1137">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="d26f3-1138">Object</span><span class="sxs-lookup"><span data-stu-id="d26f3-1138">Object</span></span>| <span data-ttu-id="d26f3-1139">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1139">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1140">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1140">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d26f3-1141">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1141">Object</span></span>| <span data-ttu-id="d26f3-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1143">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1143">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d26f3-1144">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-1144">function</span></span>||<span data-ttu-id="d26f3-1145">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-1145">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d26f3-1146">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1146">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d26f3-1147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1147">Requirements</span></span>

|<span data-ttu-id="d26f3-1148">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1148">Requirement</span></span>| <span data-ttu-id="d26f3-1149">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1149">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1151">1.3</span><span class="sxs-lookup"><span data-stu-id="d26f3-1151">1.3</span></span>|
|[<span data-ttu-id="d26f3-1152">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1153">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1153">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-1154">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-1154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1155">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-1155">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d26f3-1156">Exemplos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1156">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d26f3-p177">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d26f3-1159">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d26f3-1159">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d26f3-1160">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1160">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d26f3-p178">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d26f3-1164">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d26f3-1164">Parameters</span></span>

|<span data-ttu-id="d26f3-1165">Nome</span><span class="sxs-lookup"><span data-stu-id="d26f3-1165">Name</span></span>| <span data-ttu-id="d26f3-1166">Tipo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1166">Type</span></span>| <span data-ttu-id="d26f3-1167">Atributos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1167">Attributes</span></span>| <span data-ttu-id="d26f3-1168">Descrição</span><span class="sxs-lookup"><span data-stu-id="d26f3-1168">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d26f3-1169">String</span><span class="sxs-lookup"><span data-stu-id="d26f3-1169">String</span></span>||<span data-ttu-id="d26f3-p179">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d26f3-1173">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1173">Object</span></span>| <span data-ttu-id="d26f3-1174">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1174">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1175">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1175">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d26f3-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="d26f3-1176">Object</span></span>| <span data-ttu-id="d26f3-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1178">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1178">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d26f3-1179">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d26f3-1179">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d26f3-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d26f3-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="d26f3-1181">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1181">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d26f3-1182">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1182">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d26f3-1183">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1183">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="d26f3-1184">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1184">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d26f3-1185">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="d26f3-1185">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d26f3-1186">function</span><span class="sxs-lookup"><span data-stu-id="d26f3-1186">function</span></span>||<span data-ttu-id="d26f3-1187">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d26f3-1187">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d26f3-1188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d26f3-1188">Requirements</span></span>

|<span data-ttu-id="d26f3-1189">Requisito</span><span class="sxs-lookup"><span data-stu-id="d26f3-1189">Requirement</span></span>| <span data-ttu-id="d26f3-1190">Valor</span><span class="sxs-lookup"><span data-stu-id="d26f3-1190">Value</span></span>|
|---|---|
|[<span data-ttu-id="d26f3-1191">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d26f3-1191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d26f3-1192">1.2</span><span class="sxs-lookup"><span data-stu-id="d26f3-1192">1.2</span></span>|
|[<span data-ttu-id="d26f3-1193">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1193">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d26f3-1194">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d26f3-1194">ReadWriteItem</span></span>|
|[<span data-ttu-id="d26f3-1195">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d26f3-1195">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d26f3-1196">Escrever</span><span class="sxs-lookup"><span data-stu-id="d26f3-1196">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d26f3-1197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d26f3-1197">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
