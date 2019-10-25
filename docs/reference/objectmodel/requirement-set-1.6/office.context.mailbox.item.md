---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,6
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 570cadd13a71519dd2bf13a0f9317a2fba837a91
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37681918"
---
# <a name="item"></a><span data-ttu-id="8d9ba-102">item</span><span class="sxs-lookup"><span data-stu-id="8d9ba-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8d9ba-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8d9ba-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8d9ba-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-106">Requirements</span></span>

|<span data-ttu-id="8d9ba-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-107">Requirement</span></span>| <span data-ttu-id="8d9ba-108">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-110">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-110">1.0</span></span>|
|[<span data-ttu-id="8d9ba-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-112">Restricted</span></span>|
|[<span data-ttu-id="8d9ba-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8d9ba-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-115">Members and methods</span></span>

| <span data-ttu-id="8d9ba-116">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-116">Member</span></span> | <span data-ttu-id="8d9ba-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8d9ba-118">attachments</span><span class="sxs-lookup"><span data-stu-id="8d9ba-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="8d9ba-119">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-119">Member</span></span> |
| [<span data-ttu-id="8d9ba-120">bcc</span><span class="sxs-lookup"><span data-stu-id="8d9ba-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="8d9ba-121">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-121">Member</span></span> |
| [<span data-ttu-id="8d9ba-122">body</span><span class="sxs-lookup"><span data-stu-id="8d9ba-122">body</span></span>](#body-body) | <span data-ttu-id="8d9ba-123">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-123">Member</span></span> |
| [<span data-ttu-id="8d9ba-124">cc</span><span class="sxs-lookup"><span data-stu-id="8d9ba-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d9ba-125">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-125">Member</span></span> |
| [<span data-ttu-id="8d9ba-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="8d9ba-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8d9ba-127">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-127">Member</span></span> |
| [<span data-ttu-id="8d9ba-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8d9ba-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8d9ba-129">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-129">Member</span></span> |
| [<span data-ttu-id="8d9ba-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8d9ba-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8d9ba-131">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-131">Member</span></span> |
| [<span data-ttu-id="8d9ba-132">end</span><span class="sxs-lookup"><span data-stu-id="8d9ba-132">end</span></span>](#end-datetime) | <span data-ttu-id="8d9ba-133">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-133">Member</span></span> |
| [<span data-ttu-id="8d9ba-134">from</span><span class="sxs-lookup"><span data-stu-id="8d9ba-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="8d9ba-135">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-135">Member</span></span> |
| [<span data-ttu-id="8d9ba-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8d9ba-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8d9ba-137">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-137">Member</span></span> |
| [<span data-ttu-id="8d9ba-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="8d9ba-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8d9ba-139">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-139">Member</span></span> |
| [<span data-ttu-id="8d9ba-140">itemId</span><span class="sxs-lookup"><span data-stu-id="8d9ba-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8d9ba-141">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-141">Member</span></span> |
| [<span data-ttu-id="8d9ba-142">itemType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="8d9ba-143">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-143">Member</span></span> |
| [<span data-ttu-id="8d9ba-144">location</span><span class="sxs-lookup"><span data-stu-id="8d9ba-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="8d9ba-145">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-145">Member</span></span> |
| [<span data-ttu-id="8d9ba-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8d9ba-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8d9ba-147">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-147">Member</span></span> |
| [<span data-ttu-id="8d9ba-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8d9ba-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="8d9ba-149">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-149">Member</span></span> |
| [<span data-ttu-id="8d9ba-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8d9ba-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d9ba-151">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-151">Member</span></span> |
| [<span data-ttu-id="8d9ba-152">organizer</span><span class="sxs-lookup"><span data-stu-id="8d9ba-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="8d9ba-153">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-153">Member</span></span> |
| [<span data-ttu-id="8d9ba-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8d9ba-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d9ba-155">Member</span><span class="sxs-lookup"><span data-stu-id="8d9ba-155">Member</span></span> |
| [<span data-ttu-id="8d9ba-156">sender</span><span class="sxs-lookup"><span data-stu-id="8d9ba-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="8d9ba-157">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-157">Member</span></span> |
| [<span data-ttu-id="8d9ba-158">start</span><span class="sxs-lookup"><span data-stu-id="8d9ba-158">start</span></span>](#start-datetime) | <span data-ttu-id="8d9ba-159">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-159">Member</span></span> |
| [<span data-ttu-id="8d9ba-160">subject</span><span class="sxs-lookup"><span data-stu-id="8d9ba-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="8d9ba-161">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-161">Member</span></span> |
| [<span data-ttu-id="8d9ba-162">to</span><span class="sxs-lookup"><span data-stu-id="8d9ba-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d9ba-163">Membro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-163">Member</span></span> |
| [<span data-ttu-id="8d9ba-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8d9ba-165">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-165">Method</span></span> |
| [<span data-ttu-id="8d9ba-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8d9ba-167">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-167">Method</span></span> |
| [<span data-ttu-id="8d9ba-168">close</span><span class="sxs-lookup"><span data-stu-id="8d9ba-168">close</span></span>](#close) | <span data-ttu-id="8d9ba-169">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-169">Method</span></span> |
| [<span data-ttu-id="8d9ba-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8d9ba-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="8d9ba-171">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-171">Method</span></span> |
| [<span data-ttu-id="8d9ba-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8d9ba-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="8d9ba-173">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-173">Method</span></span> |
| [<span data-ttu-id="8d9ba-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="8d9ba-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="8d9ba-175">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-175">Method</span></span> |
| [<span data-ttu-id="8d9ba-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8d9ba-177">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-177">Method</span></span> |
| [<span data-ttu-id="8d9ba-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8d9ba-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8d9ba-179">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-179">Method</span></span> |
| [<span data-ttu-id="8d9ba-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8d9ba-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8d9ba-181">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-181">Method</span></span> |
| [<span data-ttu-id="8d9ba-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8d9ba-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8d9ba-183">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-183">Method</span></span> |
| [<span data-ttu-id="8d9ba-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8d9ba-185">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-185">Method</span></span> |
| [<span data-ttu-id="8d9ba-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8d9ba-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="8d9ba-187">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-187">Method</span></span> |
| [<span data-ttu-id="8d9ba-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8d9ba-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8d9ba-189">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-189">Method</span></span> |
| [<span data-ttu-id="8d9ba-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8d9ba-191">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-191">Method</span></span> |
| [<span data-ttu-id="8d9ba-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8d9ba-193">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-193">Method</span></span> |
| [<span data-ttu-id="8d9ba-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8d9ba-195">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-195">Method</span></span> |
| [<span data-ttu-id="8d9ba-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8d9ba-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8d9ba-197">Método</span><span class="sxs-lookup"><span data-stu-id="8d9ba-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8d9ba-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-198">Example</span></span>

<span data-ttu-id="8d9ba-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8d9ba-200">Members</span><span class="sxs-lookup"><span data-stu-id="8d9ba-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-201">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="8d9ba-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="8d9ba-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8d9ba-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-206">Type</span></span>

*   <span data-ttu-id="8d9ba-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="8d9ba-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-208">Requirements</span></span>

|<span data-ttu-id="8d9ba-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-209">Requirement</span></span>| <span data-ttu-id="8d9ba-210">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-212">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-212">1.0</span></span>|
|[<span data-ttu-id="8d9ba-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-214">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-215">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-216">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-217">Example</span></span>

<span data-ttu-id="8d9ba-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-219">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8d9ba-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-221">Compose mode only.</span></span>

<span data-ttu-id="8d9ba-222">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-223">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8d9ba-224">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="8d9ba-225">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-226">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-226">Type</span></span>

*   [<span data-ttu-id="8d9ba-227">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8d9ba-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-228">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-228">Requirements</span></span>

|<span data-ttu-id="8d9ba-229">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-229">Requirement</span></span>| <span data-ttu-id="8d9ba-230">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-231">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-232">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9ba-232">1.1</span></span>|
|[<span data-ttu-id="8d9ba-233">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-234">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-235">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-236">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-237">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="8d9ba-238">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-239">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-240">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-240">Type</span></span>

*   [<span data-ttu-id="8d9ba-241">Body</span><span class="sxs-lookup"><span data-stu-id="8d9ba-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-242">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-242">Requirements</span></span>

|<span data-ttu-id="8d9ba-243">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-243">Requirement</span></span>| <span data-ttu-id="8d9ba-244">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-245">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-246">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9ba-246">1.1</span></span>|
|[<span data-ttu-id="8d9ba-247">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-248">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-249">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-250">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-251">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-251">Example</span></span>

<span data-ttu-id="8d9ba-252">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8d9ba-253">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-254">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-255">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8d9ba-256">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-257">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-257">Read mode</span></span>

<span data-ttu-id="8d9ba-258">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="8d9ba-259">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-260">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-261">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-261">Compose mode</span></span>

<span data-ttu-id="8d9ba-262">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="8d9ba-263">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-264">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8d9ba-265">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="8d9ba-266">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-267">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-267">Type</span></span>

*   <span data-ttu-id="8d9ba-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-269">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-269">Requirements</span></span>

|<span data-ttu-id="8d9ba-270">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-270">Requirement</span></span>| <span data-ttu-id="8d9ba-271">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-272">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-273">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-273">1.0</span></span>|
|[<span data-ttu-id="8d9ba-274">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-275">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-276">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-277">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8d9ba-278">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="8d9ba-279">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8d9ba-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8d9ba-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-284">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-284">Type</span></span>

*   <span data-ttu-id="8d9ba-285">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-286">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-286">Requirements</span></span>

|<span data-ttu-id="8d9ba-287">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-287">Requirement</span></span>| <span data-ttu-id="8d9ba-288">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-289">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-290">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-290">1.0</span></span>|
|[<span data-ttu-id="8d9ba-291">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-292">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-293">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-294">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-295">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="8d9ba-296">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="8d9ba-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="8d9ba-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-299">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-299">Type</span></span>

*   <span data-ttu-id="8d9ba-300">Data</span><span class="sxs-lookup"><span data-stu-id="8d9ba-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-301">Requirements</span></span>

|<span data-ttu-id="8d9ba-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-302">Requirement</span></span>| <span data-ttu-id="8d9ba-303">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-304">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-305">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-305">1.0</span></span>|
|[<span data-ttu-id="8d9ba-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-307">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-309">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="8d9ba-311">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="8d9ba-311">dateTimeModified: Date</span></span>

<span data-ttu-id="8d9ba-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-314">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-315">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-315">Type</span></span>

*   <span data-ttu-id="8d9ba-316">Data</span><span class="sxs-lookup"><span data-stu-id="8d9ba-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-317">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-317">Requirements</span></span>

|<span data-ttu-id="8d9ba-318">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-318">Requirement</span></span>| <span data-ttu-id="8d9ba-319">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-320">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-321">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-321">1.0</span></span>|
|[<span data-ttu-id="8d9ba-322">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-323">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-324">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-325">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-326">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="8d9ba-327">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-328">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8d9ba-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-331">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-331">Read mode</span></span>

<span data-ttu-id="8d9ba-332">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-333">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-333">Compose mode</span></span>

<span data-ttu-id="8d9ba-334">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8d9ba-335">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8d9ba-336">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8d9ba-337">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-337">Type</span></span>

*   <span data-ttu-id="8d9ba-338">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-339">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-339">Requirements</span></span>

|<span data-ttu-id="8d9ba-340">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-340">Requirement</span></span>| <span data-ttu-id="8d9ba-341">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-342">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-343">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-343">1.0</span></span>|
|[<span data-ttu-id="8d9ba-344">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-345">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-346">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-347">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-348">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-p114">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8d9ba-p115">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-353">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-354">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-354">Type</span></span>

*   [<span data-ttu-id="8d9ba-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d9ba-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="8d9ba-356">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="8d9ba-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-357">Requirements</span></span>

|<span data-ttu-id="8d9ba-358">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-358">Requirement</span></span>| <span data-ttu-id="8d9ba-359">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-361">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-361">1.0</span></span>|
|[<span data-ttu-id="8d9ba-362">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-363">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-364">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-365">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="8d9ba-366">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-366">internetMessageId: String</span></span>

<span data-ttu-id="8d9ba-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-369">Type</span></span>

*   <span data-ttu-id="8d9ba-370">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-371">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-371">Requirements</span></span>

|<span data-ttu-id="8d9ba-372">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-372">Requirement</span></span>| <span data-ttu-id="8d9ba-373">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-374">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-375">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-375">1.0</span></span>|
|[<span data-ttu-id="8d9ba-376">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-377">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-378">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-379">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-380">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="8d9ba-381">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="8d9ba-381">itemClass: String</span></span>

<span data-ttu-id="8d9ba-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8d9ba-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8d9ba-386">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-386">Type</span></span> | <span data-ttu-id="8d9ba-387">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-387">Description</span></span> | <span data-ttu-id="8d9ba-388">classe de item</span><span class="sxs-lookup"><span data-stu-id="8d9ba-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8d9ba-389">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="8d9ba-389">Appointment items</span></span> | <span data-ttu-id="8d9ba-390">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8d9ba-391">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-391">Message items</span></span> | <span data-ttu-id="8d9ba-392">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8d9ba-393">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-394">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-394">Type</span></span>

*   <span data-ttu-id="8d9ba-395">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-396">Requirements</span></span>

|<span data-ttu-id="8d9ba-397">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-397">Requirement</span></span>| <span data-ttu-id="8d9ba-398">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-399">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-400">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-400">1.0</span></span>|
|[<span data-ttu-id="8d9ba-401">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-402">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-403">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-404">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-405">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8d9ba-406">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-406">(nullable) itemId: String</span></span>

<span data-ttu-id="8d9ba-p119">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-409">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8d9ba-410">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8d9ba-411">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8d9ba-412">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8d9ba-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-415">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-415">Type</span></span>

*   <span data-ttu-id="8d9ba-416">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-417">Requirements</span></span>

|<span data-ttu-id="8d9ba-418">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-418">Requirement</span></span>| <span data-ttu-id="8d9ba-419">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-420">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-421">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-421">1.0</span></span>|
|[<span data-ttu-id="8d9ba-422">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-423">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-424">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-425">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-426">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-426">Example</span></span>

<span data-ttu-id="8d9ba-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="8d9ba-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-430">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8d9ba-431">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-432">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-432">Type</span></span>

*   [<span data-ttu-id="8d9ba-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-434">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-434">Requirements</span></span>

|<span data-ttu-id="8d9ba-435">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-435">Requirement</span></span>| <span data-ttu-id="8d9ba-436">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-437">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-438">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-438">1.0</span></span>|
|[<span data-ttu-id="8d9ba-439">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-440">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-441">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-442">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-443">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="8d9ba-444">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-445">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-446">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-446">Read mode</span></span>

<span data-ttu-id="8d9ba-447">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-448">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-448">Compose mode</span></span>

<span data-ttu-id="8d9ba-449">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-450">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-450">Type</span></span>

*   <span data-ttu-id="8d9ba-451">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-452">Requirements</span></span>

|<span data-ttu-id="8d9ba-453">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-453">Requirement</span></span>| <span data-ttu-id="8d9ba-454">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-455">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-456">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-456">1.0</span></span>|
|[<span data-ttu-id="8d9ba-457">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-458">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-459">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-460">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8d9ba-461">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-461">normalizedSubject: String</span></span>

<span data-ttu-id="8d9ba-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8d9ba-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-466">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-466">Type</span></span>

*   <span data-ttu-id="8d9ba-467">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-468">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-468">Requirements</span></span>

|<span data-ttu-id="8d9ba-469">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-469">Requirement</span></span>| <span data-ttu-id="8d9ba-470">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-471">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-472">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-472">1.0</span></span>|
|[<span data-ttu-id="8d9ba-473">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-474">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-475">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-476">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-477">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="8d9ba-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-479">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-480">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-480">Type</span></span>

*   [<span data-ttu-id="8d9ba-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8d9ba-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-482">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-482">Requirements</span></span>

|<span data-ttu-id="8d9ba-483">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-483">Requirement</span></span>| <span data-ttu-id="8d9ba-484">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-485">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-486">1.3</span><span class="sxs-lookup"><span data-stu-id="8d9ba-486">1.3</span></span>|
|[<span data-ttu-id="8d9ba-487">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-488">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-489">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-490">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-491">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-492">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-493">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8d9ba-494">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-495">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-495">Read mode</span></span>

<span data-ttu-id="8d9ba-496">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="8d9ba-497">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-498">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-499">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-499">Compose mode</span></span>

<span data-ttu-id="8d9ba-500">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="8d9ba-501">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-502">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8d9ba-503">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="8d9ba-504">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-505">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-505">Type</span></span>

*   <span data-ttu-id="8d9ba-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-507">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-507">Requirements</span></span>

|<span data-ttu-id="8d9ba-508">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-508">Requirement</span></span>| <span data-ttu-id="8d9ba-509">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-510">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-511">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-511">1.0</span></span>|
|[<span data-ttu-id="8d9ba-512">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-513">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-514">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-515">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-516">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-p128">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-519">Type</span></span>

*   [<span data-ttu-id="8d9ba-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d9ba-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-521">Requirements</span></span>

|<span data-ttu-id="8d9ba-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-522">Requirement</span></span>| <span data-ttu-id="8d9ba-523">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-525">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-525">1.0</span></span>|
|[<span data-ttu-id="8d9ba-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-527">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-528">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-529">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-530">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-531">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-532">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8d9ba-533">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-534">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-534">Read mode</span></span>

<span data-ttu-id="8d9ba-535">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="8d9ba-536">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-537">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-538">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-538">Compose mode</span></span>

<span data-ttu-id="8d9ba-539">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="8d9ba-540">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-541">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8d9ba-542">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="8d9ba-543">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-544">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-544">Type</span></span>

*   <span data-ttu-id="8d9ba-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-546">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-546">Requirements</span></span>

|<span data-ttu-id="8d9ba-547">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-547">Requirement</span></span>| <span data-ttu-id="8d9ba-548">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-549">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-550">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-550">1.0</span></span>|
|[<span data-ttu-id="8d9ba-551">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-552">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-553">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-554">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-555">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-p132">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8d9ba-p133">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-560">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d9ba-561">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-561">Type</span></span>

*   [<span data-ttu-id="8d9ba-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d9ba-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="8d9ba-563">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-563">Requirements</span></span>

|<span data-ttu-id="8d9ba-564">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-564">Requirement</span></span>| <span data-ttu-id="8d9ba-565">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-566">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-567">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-567">1.0</span></span>|
|[<span data-ttu-id="8d9ba-568">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-569">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-570">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-571">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-572">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="8d9ba-573">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-574">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8d9ba-p134">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-577">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-577">Read mode</span></span>

<span data-ttu-id="8d9ba-578">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-579">Compose mode</span></span>

<span data-ttu-id="8d9ba-580">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8d9ba-581">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8d9ba-582">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8d9ba-583">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-583">Type</span></span>

*   <span data-ttu-id="8d9ba-584">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-585">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-585">Requirements</span></span>

|<span data-ttu-id="8d9ba-586">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-586">Requirement</span></span>| <span data-ttu-id="8d9ba-587">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-588">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-589">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-589">1.0</span></span>|
|[<span data-ttu-id="8d9ba-590">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-591">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-592">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-593">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="8d9ba-594">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-595">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8d9ba-596">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-597">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-597">Read mode</span></span>

<span data-ttu-id="8d9ba-p135">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-600">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-600">Compose mode</span></span>

<span data-ttu-id="8d9ba-601">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-602">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-602">Type</span></span>

*   <span data-ttu-id="8d9ba-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-604">Requirements</span></span>

|<span data-ttu-id="8d9ba-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-605">Requirement</span></span>| <span data-ttu-id="8d9ba-606">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-607">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-608">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-608">1.0</span></span>|
|[<span data-ttu-id="8d9ba-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-610">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-611">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-612">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="8d9ba-613">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="8d9ba-614">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8d9ba-615">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d9ba-616">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8d9ba-616">Read mode</span></span>

<span data-ttu-id="8d9ba-617">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="8d9ba-618">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-619">No entanto, no Windows e no Mac, você pode configurar para obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d9ba-620">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8d9ba-620">Compose mode</span></span>

<span data-ttu-id="8d9ba-621">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="8d9ba-622">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8d9ba-623">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8d9ba-624">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="8d9ba-625">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d9ba-626">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-626">Type</span></span>

*   <span data-ttu-id="8d9ba-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-628">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-628">Requirements</span></span>

|<span data-ttu-id="8d9ba-629">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-629">Requirement</span></span>| <span data-ttu-id="8d9ba-630">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-631">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-632">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-632">1.0</span></span>|
|[<span data-ttu-id="8d9ba-633">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-634">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-635">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-636">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8d9ba-637">Métodos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8d9ba-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8d9ba-639">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8d9ba-640">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8d9ba-641">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-642">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-642">Parameters</span></span>

|<span data-ttu-id="8d9ba-643">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-643">Name</span></span>| <span data-ttu-id="8d9ba-644">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-644">Type</span></span>| <span data-ttu-id="8d9ba-645">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-645">Attributes</span></span>| <span data-ttu-id="8d9ba-646">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8d9ba-647">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-647">String</span></span>||<span data-ttu-id="8d9ba-p139">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8d9ba-650">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-650">String</span></span>||<span data-ttu-id="8d9ba-p140">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8d9ba-653">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-653">Object</span></span>| <span data-ttu-id="8d9ba-654">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-654">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-655">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="8d9ba-656">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-656">Object</span></span> | <span data-ttu-id="8d9ba-657">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-657">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-658">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="8d9ba-659">Booliano</span><span class="sxs-lookup"><span data-stu-id="8d9ba-659">Boolean</span></span> | <span data-ttu-id="8d9ba-660">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-660">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-661">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="8d9ba-662">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-662">function</span></span>| <span data-ttu-id="8d9ba-663">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-663">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-664">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d9ba-665">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8d9ba-666">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d9ba-667">Erros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-667">Errors</span></span>

| <span data-ttu-id="8d9ba-668">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-668">Error code</span></span> | <span data-ttu-id="8d9ba-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8d9ba-670">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8d9ba-671">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8d9ba-672">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-673">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-673">Requirements</span></span>

|<span data-ttu-id="8d9ba-674">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-674">Requirement</span></span>| <span data-ttu-id="8d9ba-675">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-676">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-677">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9ba-677">1.1</span></span>|
|[<span data-ttu-id="8d9ba-678">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d9ba-680">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-681">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d9ba-682">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-682">Examples</span></span>

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

<span data-ttu-id="8d9ba-683">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8d9ba-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8d9ba-685">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8d9ba-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8d9ba-689">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8d9ba-690">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-691">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-691">Parameters</span></span>

|<span data-ttu-id="8d9ba-692">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-692">Name</span></span>| <span data-ttu-id="8d9ba-693">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-693">Type</span></span>| <span data-ttu-id="8d9ba-694">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-694">Attributes</span></span>| <span data-ttu-id="8d9ba-695">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8d9ba-696">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-696">String</span></span>||<span data-ttu-id="8d9ba-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8d9ba-699">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-699">String</span></span>||<span data-ttu-id="8d9ba-700">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-700">The subject of the item to be attached.</span></span> <span data-ttu-id="8d9ba-701">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8d9ba-702">Object</span><span class="sxs-lookup"><span data-stu-id="8d9ba-702">Object</span></span>| <span data-ttu-id="8d9ba-703">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-703">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-704">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d9ba-705">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-705">Object</span></span>| <span data-ttu-id="8d9ba-706">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-706">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-707">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d9ba-708">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-708">function</span></span>| <span data-ttu-id="8d9ba-709">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-709">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-710">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d9ba-711">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8d9ba-712">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d9ba-713">Erros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-713">Errors</span></span>

| <span data-ttu-id="8d9ba-714">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-714">Error code</span></span> | <span data-ttu-id="8d9ba-715">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8d9ba-716">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-717">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-717">Requirements</span></span>

|<span data-ttu-id="8d9ba-718">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-718">Requirement</span></span>| <span data-ttu-id="8d9ba-719">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-720">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-721">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9ba-721">1.1</span></span>|
|[<span data-ttu-id="8d9ba-722">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d9ba-724">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-725">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-726">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-726">Example</span></span>

<span data-ttu-id="8d9ba-727">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="8d9ba-728">close()</span><span class="sxs-lookup"><span data-stu-id="8d9ba-728">close()</span></span>

<span data-ttu-id="8d9ba-729">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8d9ba-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-732">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8d9ba-733">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-734">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-734">Requirements</span></span>

|<span data-ttu-id="8d9ba-735">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-735">Requirement</span></span>| <span data-ttu-id="8d9ba-736">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-737">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-738">1.3</span><span class="sxs-lookup"><span data-stu-id="8d9ba-738">1.3</span></span>|
|[<span data-ttu-id="8d9ba-739">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-740">Restrito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-740">Restricted</span></span>|
|[<span data-ttu-id="8d9ba-741">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-742">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8d9ba-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8d9ba-744">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-745">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-746">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8d9ba-747">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8d9ba-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-751">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-751">Parameters</span></span>

| <span data-ttu-id="8d9ba-752">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-752">Name</span></span> | <span data-ttu-id="8d9ba-753">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-753">Type</span></span> | <span data-ttu-id="8d9ba-754">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-754">Attributes</span></span> | <span data-ttu-id="8d9ba-755">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8d9ba-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8d9ba-756">String &#124; Object</span></span>| |<span data-ttu-id="8d9ba-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8d9ba-759">**OU**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-759">**OR**</span></span><br/><span data-ttu-id="8d9ba-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8d9ba-762">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-762">String</span></span> | <span data-ttu-id="8d9ba-763">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-763">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8d9ba-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8d9ba-767">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-767">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-768">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8d9ba-769">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-769">String</span></span> | | <span data-ttu-id="8d9ba-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8d9ba-772">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-772">String</span></span> | | <span data-ttu-id="8d9ba-773">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8d9ba-774">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-774">String</span></span> | | <span data-ttu-id="8d9ba-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8d9ba-777">Booliano</span><span class="sxs-lookup"><span data-stu-id="8d9ba-777">Boolean</span></span> | | <span data-ttu-id="8d9ba-p151">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8d9ba-780">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-780">String</span></span> | | <span data-ttu-id="8d9ba-p152">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8d9ba-784">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-784">function</span></span> | <span data-ttu-id="8d9ba-785">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-785">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-786">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-787">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-787">Requirements</span></span>

|<span data-ttu-id="8d9ba-788">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-788">Requirement</span></span>| <span data-ttu-id="8d9ba-789">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-790">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-791">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-791">1.0</span></span>|
|[<span data-ttu-id="8d9ba-792">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-793">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-794">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-795">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d9ba-796">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-796">Examples</span></span>

<span data-ttu-id="8d9ba-797">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8d9ba-798">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8d9ba-799">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8d9ba-800">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8d9ba-801">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8d9ba-802">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8d9ba-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8d9ba-804">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-805">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-806">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8d9ba-807">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8d9ba-p153">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-811">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-811">Parameters</span></span>

| <span data-ttu-id="8d9ba-812">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-812">Name</span></span> | <span data-ttu-id="8d9ba-813">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-813">Type</span></span> | <span data-ttu-id="8d9ba-814">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-814">Attributes</span></span> | <span data-ttu-id="8d9ba-815">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8d9ba-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8d9ba-816">String &#124; Object</span></span>| | <span data-ttu-id="8d9ba-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8d9ba-819">**OU**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-819">**OR**</span></span><br/><span data-ttu-id="8d9ba-p155">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8d9ba-822">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-822">String</span></span> | <span data-ttu-id="8d9ba-823">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-823">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-p156">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8d9ba-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8d9ba-827">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-827">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-828">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8d9ba-829">Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-829">String</span></span> | | <span data-ttu-id="8d9ba-p157">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8d9ba-832">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-832">String</span></span> | | <span data-ttu-id="8d9ba-833">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8d9ba-834">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-834">String</span></span> | | <span data-ttu-id="8d9ba-p158">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8d9ba-837">Booliano</span><span class="sxs-lookup"><span data-stu-id="8d9ba-837">Boolean</span></span> | | <span data-ttu-id="8d9ba-p159">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8d9ba-840">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-840">String</span></span> | | <span data-ttu-id="8d9ba-p160">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8d9ba-844">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-844">function</span></span> | <span data-ttu-id="8d9ba-845">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-845">&lt;optional&gt;</span></span> | <span data-ttu-id="8d9ba-846">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-847">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-847">Requirements</span></span>

|<span data-ttu-id="8d9ba-848">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-848">Requirement</span></span>| <span data-ttu-id="8d9ba-849">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-850">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-851">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-851">1.0</span></span>|
|[<span data-ttu-id="8d9ba-852">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-853">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-854">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-855">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d9ba-856">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-856">Examples</span></span>

<span data-ttu-id="8d9ba-857">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8d9ba-858">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8d9ba-859">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8d9ba-860">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8d9ba-861">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8d9ba-862">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="8d9ba-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="8d9ba-864">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-865">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-866">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-866">Requirements</span></span>

|<span data-ttu-id="8d9ba-867">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-867">Requirement</span></span>| <span data-ttu-id="8d9ba-868">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-869">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-870">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-870">1.0</span></span>|
|[<span data-ttu-id="8d9ba-871">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-872">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-873">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-874">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-875">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-875">Returns:</span></span>

<span data-ttu-id="8d9ba-876">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-877">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-877">Example</span></span>

<span data-ttu-id="8d9ba-878">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="8d9ba-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="8d9ba-880">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-881">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-882">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-882">Parameters</span></span>

|<span data-ttu-id="8d9ba-883">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-883">Name</span></span>| <span data-ttu-id="8d9ba-884">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-884">Type</span></span>| <span data-ttu-id="8d9ba-885">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8d9ba-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="8d9ba-887">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-888">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-888">Requirements</span></span>

|<span data-ttu-id="8d9ba-889">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-889">Requirement</span></span>| <span data-ttu-id="8d9ba-890">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-891">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-892">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-892">1.0</span></span>|
|[<span data-ttu-id="8d9ba-893">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-894">Restrito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-894">Restricted</span></span>|
|[<span data-ttu-id="8d9ba-895">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-896">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-897">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-897">Returns:</span></span>

<span data-ttu-id="8d9ba-898">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8d9ba-899">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8d9ba-900">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8d9ba-901">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8d9ba-902">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8d9ba-902">Value of `entityType`</span></span> | <span data-ttu-id="8d9ba-903">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="8d9ba-903">Type of objects in returned array</span></span> | <span data-ttu-id="8d9ba-904">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="8d9ba-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8d9ba-905">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-905">String</span></span> | <span data-ttu-id="8d9ba-906">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8d9ba-907">Contato</span><span class="sxs-lookup"><span data-stu-id="8d9ba-907">Contact</span></span> | <span data-ttu-id="8d9ba-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8d9ba-909">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-909">String</span></span> | <span data-ttu-id="8d9ba-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8d9ba-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8d9ba-911">MeetingSuggestion</span></span> | <span data-ttu-id="8d9ba-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8d9ba-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8d9ba-913">PhoneNumber</span></span> | <span data-ttu-id="8d9ba-914">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8d9ba-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8d9ba-915">TaskSuggestion</span></span> | <span data-ttu-id="8d9ba-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8d9ba-917">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-917">String</span></span> | <span data-ttu-id="8d9ba-918">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8d9ba-918">**Restricted**</span></span> |

<span data-ttu-id="8d9ba-919">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="8d9ba-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-920">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-920">Example</span></span>

<span data-ttu-id="8d9ba-921">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="8d9ba-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="8d9ba-923">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-924">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-925">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-926">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-926">Parameters</span></span>

|<span data-ttu-id="8d9ba-927">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-927">Name</span></span>| <span data-ttu-id="8d9ba-928">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-928">Type</span></span>| <span data-ttu-id="8d9ba-929">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8d9ba-930">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-930">String</span></span>|<span data-ttu-id="8d9ba-931">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-932">Requirements</span></span>

|<span data-ttu-id="8d9ba-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-933">Requirement</span></span>| <span data-ttu-id="8d9ba-934">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-935">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-936">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-936">1.0</span></span>|
|[<span data-ttu-id="8d9ba-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-938">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-939">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-940">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-941">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-941">Returns:</span></span>

<span data-ttu-id="8d9ba-p162">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8d9ba-944">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="8d9ba-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="8d9ba-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8d9ba-946">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-947">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-p163">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8d9ba-951">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8d9ba-952">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8d9ba-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-956">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-956">Requirements</span></span>

|<span data-ttu-id="8d9ba-957">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-957">Requirement</span></span>| <span data-ttu-id="8d9ba-958">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-959">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-960">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-960">1.0</span></span>|
|[<span data-ttu-id="8d9ba-961">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-962">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-963">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-964">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-965">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-965">Returns:</span></span>

<span data-ttu-id="8d9ba-p165">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="8d9ba-968">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-969">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-969">Example</span></span>

<span data-ttu-id="8d9ba-970">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8d9ba-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8d9ba-972">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-973">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-974">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8d9ba-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-977">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-977">Parameters</span></span>

|<span data-ttu-id="8d9ba-978">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-978">Name</span></span>| <span data-ttu-id="8d9ba-979">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-979">Type</span></span>| <span data-ttu-id="8d9ba-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8d9ba-981">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-981">String</span></span>|<span data-ttu-id="8d9ba-982">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-983">Requirements</span></span>

|<span data-ttu-id="8d9ba-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-984">Requirement</span></span>| <span data-ttu-id="8d9ba-985">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-986">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-987">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-987">1.0</span></span>|
|[<span data-ttu-id="8d9ba-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-989">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-990">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-991">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-992">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-992">Returns:</span></span>

<span data-ttu-id="8d9ba-993">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="8d9ba-994">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8d9ba-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-995">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8d9ba-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8d9ba-997">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8d9ba-p167">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-1000">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1000">Parameters</span></span>

|<span data-ttu-id="8d9ba-1001">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1001">Name</span></span>| <span data-ttu-id="8d9ba-1002">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1002">Type</span></span>| <span data-ttu-id="8d9ba-1003">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1003">Attributes</span></span>| <span data-ttu-id="8d9ba-1004">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1004">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8d9ba-1005">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1005">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8d9ba-p168">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8d9ba-1009">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1009">Object</span></span>| <span data-ttu-id="8d9ba-1010">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1011">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1011">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d9ba-1012">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1012">Object</span></span>| <span data-ttu-id="8d9ba-1013">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1014">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1014">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d9ba-1015">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1015">function</span></span>||<span data-ttu-id="8d9ba-1016">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8d9ba-1017">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1017">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8d9ba-1018">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1018">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1019">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1019">Requirements</span></span>

|<span data-ttu-id="8d9ba-1020">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1020">Requirement</span></span>| <span data-ttu-id="8d9ba-1021">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1021">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1022">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1022">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1023">1.2</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1023">1.2</span></span>|
|[<span data-ttu-id="8d9ba-1024">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1024">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1025">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1025">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-1026">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1026">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1027">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1027">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-1028">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1028">Returns:</span></span>

<span data-ttu-id="8d9ba-1029">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1029">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="8d9ba-1030">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1030">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-1031">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1031">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="8d9ba-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="8d9ba-1033">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1033">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="8d9ba-1034">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1034">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-1035">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1035">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1036">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1036">Requirements</span></span>

|<span data-ttu-id="8d9ba-1037">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1037">Requirement</span></span>| <span data-ttu-id="8d9ba-1038">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1038">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1039">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1039">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1040">1.6</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1040">1.6</span></span> |
|[<span data-ttu-id="8d9ba-1041">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1041">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1042">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1042">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-1043">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1043">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1044">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1044">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-1045">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1045">Returns:</span></span>

<span data-ttu-id="8d9ba-1046">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1046">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-1047">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1047">Example</span></span>

<span data-ttu-id="8d9ba-1048">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1048">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8d9ba-1049">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1049">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8d9ba-p171">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-1052">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1052">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d9ba-p172">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8d9ba-1056">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1056">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8d9ba-1057">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1057">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8d9ba-p173">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1061">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1061">Requirements</span></span>

|<span data-ttu-id="8d9ba-1062">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1062">Requirement</span></span>| <span data-ttu-id="8d9ba-1063">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1063">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1064">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1064">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1065">1.6</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1065">1.6</span></span> |
|[<span data-ttu-id="8d9ba-1066">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1066">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1067">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1067">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-1068">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1068">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1069">Read</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1069">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d9ba-1070">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1070">Returns:</span></span>

<span data-ttu-id="8d9ba-p174">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8d9ba-1073">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1073">Example</span></span>

<span data-ttu-id="8d9ba-1074">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1074">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8d9ba-1075">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1075">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8d9ba-1076">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1076">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8d9ba-p175">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p175">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-1080">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1080">Parameters</span></span>

|<span data-ttu-id="8d9ba-1081">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1081">Name</span></span>| <span data-ttu-id="8d9ba-1082">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1082">Type</span></span>| <span data-ttu-id="8d9ba-1083">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1083">Attributes</span></span>| <span data-ttu-id="8d9ba-1084">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1084">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8d9ba-1085">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1085">function</span></span>||<span data-ttu-id="8d9ba-1086">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8d9ba-1087">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1087">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8d9ba-1088">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1088">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8d9ba-1089">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1089">Object</span></span>| <span data-ttu-id="8d9ba-1090">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1091">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1091">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8d9ba-1092">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1092">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1093">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1093">Requirements</span></span>

|<span data-ttu-id="8d9ba-1094">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1094">Requirement</span></span>| <span data-ttu-id="8d9ba-1095">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1096">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1097">1.0</span></span>|
|[<span data-ttu-id="8d9ba-1098">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1099">ReadItem</span></span>|
|[<span data-ttu-id="8d9ba-1100">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1101">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1101">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-1102">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1102">Example</span></span>

<span data-ttu-id="8d9ba-p178">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p178">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8d9ba-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8d9ba-1107">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1107">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8d9ba-1108">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1108">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8d9ba-1109">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1109">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8d9ba-1110">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8d9ba-1111">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1111">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-1112">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1112">Parameters</span></span>

|<span data-ttu-id="8d9ba-1113">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1113">Name</span></span>| <span data-ttu-id="8d9ba-1114">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1114">Type</span></span>| <span data-ttu-id="8d9ba-1115">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1115">Attributes</span></span>| <span data-ttu-id="8d9ba-1116">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8d9ba-1117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1117">String</span></span>||<span data-ttu-id="8d9ba-1118">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1118">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8d9ba-1119">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1119">Object</span></span>| <span data-ttu-id="8d9ba-1120">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1121">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d9ba-1122">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1122">Object</span></span>| <span data-ttu-id="8d9ba-1123">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1124">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d9ba-1125">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1125">function</span></span>| <span data-ttu-id="8d9ba-1126">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1127">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d9ba-1128">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1128">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d9ba-1129">Erros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1129">Errors</span></span>

| <span data-ttu-id="8d9ba-1130">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1130">Error code</span></span> | <span data-ttu-id="8d9ba-1131">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1131">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8d9ba-1132">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1132">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1133">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1133">Requirements</span></span>

|<span data-ttu-id="8d9ba-1134">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1134">Requirement</span></span>| <span data-ttu-id="8d9ba-1135">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1136">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1137">1.1</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1137">1.1</span></span>|
|[<span data-ttu-id="8d9ba-1138">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1139">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1139">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d9ba-1140">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1141">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1141">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-1142">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1142">Example</span></span>

<span data-ttu-id="8d9ba-1143">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1143">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="8d9ba-1144">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1144">saveAsync([options], callback)</span></span>

<span data-ttu-id="8d9ba-1145">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1145">Asynchronously saves an item.</span></span>

<span data-ttu-id="8d9ba-1146">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1146">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="8d9ba-1147">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1147">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="8d9ba-1148">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1148">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-1149">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1149">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="8d9ba-1150">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1150">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8d9ba-p182">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8d9ba-1154">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1154">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8d9ba-1155">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1155">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="8d9ba-1156">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1156">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="8d9ba-1157">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1157">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="8d9ba-1158">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1158">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-1159">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1159">Parameters</span></span>

|<span data-ttu-id="8d9ba-1160">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1160">Name</span></span>| <span data-ttu-id="8d9ba-1161">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1161">Type</span></span>| <span data-ttu-id="8d9ba-1162">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1162">Attributes</span></span>| <span data-ttu-id="8d9ba-1163">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1163">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="8d9ba-1164">Object</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1164">Object</span></span>| <span data-ttu-id="8d9ba-1165">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1165">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1166">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1166">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d9ba-1167">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1167">Object</span></span>| <span data-ttu-id="8d9ba-1168">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1168">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1169">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1169">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d9ba-1170">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1170">function</span></span>||<span data-ttu-id="8d9ba-1171">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1171">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8d9ba-1172">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1172">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1173">Requirements</span></span>

|<span data-ttu-id="8d9ba-1174">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1174">Requirement</span></span>| <span data-ttu-id="8d9ba-1175">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1177">1.3</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1177">1.3</span></span>|
|[<span data-ttu-id="8d9ba-1178">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1179">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1179">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d9ba-1180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1181">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1181">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d9ba-1182">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1182">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="8d9ba-p184">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8d9ba-1185">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1185">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8d9ba-1186">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1186">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8d9ba-p185">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d9ba-1190">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1190">Parameters</span></span>

|<span data-ttu-id="8d9ba-1191">Nome</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1191">Name</span></span>| <span data-ttu-id="8d9ba-1192">Tipo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1192">Type</span></span>| <span data-ttu-id="8d9ba-1193">Atributos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1193">Attributes</span></span>| <span data-ttu-id="8d9ba-1194">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1194">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8d9ba-1195">String</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1195">String</span></span>||<span data-ttu-id="8d9ba-p186">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8d9ba-1199">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1199">Object</span></span>| <span data-ttu-id="8d9ba-1200">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1201">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1201">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d9ba-1202">Objeto</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1202">Object</span></span>| <span data-ttu-id="8d9ba-1203">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1203">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1204">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1204">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8d9ba-1205">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1205">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8d9ba-1206">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1206">&lt;optional&gt;</span></span>|<span data-ttu-id="8d9ba-1207">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1207">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="8d9ba-1208">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1208">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8d9ba-1209">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1209">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="8d9ba-1210">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1210">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8d9ba-1211">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1211">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8d9ba-1212">function</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1212">function</span></span>||<span data-ttu-id="8d9ba-1213">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d9ba-1214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1214">Requirements</span></span>

|<span data-ttu-id="8d9ba-1215">Requisito</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1215">Requirement</span></span>| <span data-ttu-id="8d9ba-1216">Valor</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1216">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d9ba-1217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d9ba-1218">1.2</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1218">1.2</span></span>|
|[<span data-ttu-id="8d9ba-1219">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1219">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d9ba-1220">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1220">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d9ba-1221">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1221">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d9ba-1222">Escrever</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1222">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d9ba-1223">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8d9ba-1223">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
