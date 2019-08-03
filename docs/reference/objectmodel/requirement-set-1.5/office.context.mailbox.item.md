---
title: Office.context.mailbox.item - conjunto de requisitos 1.5
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: cb841795769e19b76bda49c4a9b704734130eadc
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064631"
---
# <a name="item"></a><span data-ttu-id="24966-102">item</span><span class="sxs-lookup"><span data-stu-id="24966-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="24966-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="24966-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="24966-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="24966-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-106">Requirements</span></span>

|<span data-ttu-id="24966-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-107">Requirement</span></span>| <span data-ttu-id="24966-108">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-110">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-110">1.0</span></span>|
|[<span data-ttu-id="24966-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="24966-112">Restricted</span></span>|
|[<span data-ttu-id="24966-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="24966-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="24966-115">Members and methods</span></span>

| <span data-ttu-id="24966-116">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-116">Member</span></span> | <span data-ttu-id="24966-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="24966-118">attachments</span><span class="sxs-lookup"><span data-stu-id="24966-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="24966-119">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-119">Member</span></span> |
| [<span data-ttu-id="24966-120">bcc</span><span class="sxs-lookup"><span data-stu-id="24966-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="24966-121">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-121">Member</span></span> |
| [<span data-ttu-id="24966-122">body</span><span class="sxs-lookup"><span data-stu-id="24966-122">body</span></span>](#body-body) | <span data-ttu-id="24966-123">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-123">Member</span></span> |
| [<span data-ttu-id="24966-124">cc</span><span class="sxs-lookup"><span data-stu-id="24966-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="24966-125">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-125">Member</span></span> |
| [<span data-ttu-id="24966-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="24966-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="24966-127">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-127">Member</span></span> |
| [<span data-ttu-id="24966-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="24966-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="24966-129">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-129">Member</span></span> |
| [<span data-ttu-id="24966-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="24966-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="24966-131">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-131">Member</span></span> |
| [<span data-ttu-id="24966-132">end</span><span class="sxs-lookup"><span data-stu-id="24966-132">end</span></span>](#end-datetime) | <span data-ttu-id="24966-133">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-133">Member</span></span> |
| [<span data-ttu-id="24966-134">from</span><span class="sxs-lookup"><span data-stu-id="24966-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="24966-135">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-135">Member</span></span> |
| [<span data-ttu-id="24966-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="24966-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="24966-137">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-137">Member</span></span> |
| [<span data-ttu-id="24966-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="24966-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="24966-139">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-139">Member</span></span> |
| [<span data-ttu-id="24966-140">itemId</span><span class="sxs-lookup"><span data-stu-id="24966-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="24966-141">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-141">Member</span></span> |
| [<span data-ttu-id="24966-142">itemType</span><span class="sxs-lookup"><span data-stu-id="24966-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="24966-143">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-143">Member</span></span> |
| [<span data-ttu-id="24966-144">location</span><span class="sxs-lookup"><span data-stu-id="24966-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="24966-145">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-145">Member</span></span> |
| [<span data-ttu-id="24966-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="24966-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="24966-147">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-147">Member</span></span> |
| [<span data-ttu-id="24966-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="24966-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="24966-149">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-149">Member</span></span> |
| [<span data-ttu-id="24966-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="24966-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="24966-151">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-151">Member</span></span> |
| [<span data-ttu-id="24966-152">organizer</span><span class="sxs-lookup"><span data-stu-id="24966-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="24966-153">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-153">Member</span></span> |
| [<span data-ttu-id="24966-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="24966-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="24966-155">Member</span><span class="sxs-lookup"><span data-stu-id="24966-155">Member</span></span> |
| [<span data-ttu-id="24966-156">sender</span><span class="sxs-lookup"><span data-stu-id="24966-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="24966-157">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-157">Member</span></span> |
| [<span data-ttu-id="24966-158">start</span><span class="sxs-lookup"><span data-stu-id="24966-158">start</span></span>](#start-datetime) | <span data-ttu-id="24966-159">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-159">Member</span></span> |
| [<span data-ttu-id="24966-160">subject</span><span class="sxs-lookup"><span data-stu-id="24966-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="24966-161">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-161">Member</span></span> |
| [<span data-ttu-id="24966-162">to</span><span class="sxs-lookup"><span data-stu-id="24966-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="24966-163">Membro</span><span class="sxs-lookup"><span data-stu-id="24966-163">Member</span></span> |
| [<span data-ttu-id="24966-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="24966-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="24966-165">Método</span><span class="sxs-lookup"><span data-stu-id="24966-165">Method</span></span> |
| [<span data-ttu-id="24966-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="24966-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="24966-167">Método</span><span class="sxs-lookup"><span data-stu-id="24966-167">Method</span></span> |
| [<span data-ttu-id="24966-168">close</span><span class="sxs-lookup"><span data-stu-id="24966-168">close</span></span>](#close) | <span data-ttu-id="24966-169">Método</span><span class="sxs-lookup"><span data-stu-id="24966-169">Method</span></span> |
| [<span data-ttu-id="24966-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="24966-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="24966-171">Método</span><span class="sxs-lookup"><span data-stu-id="24966-171">Method</span></span> |
| [<span data-ttu-id="24966-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="24966-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="24966-173">Método</span><span class="sxs-lookup"><span data-stu-id="24966-173">Method</span></span> |
| [<span data-ttu-id="24966-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="24966-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="24966-175">Método</span><span class="sxs-lookup"><span data-stu-id="24966-175">Method</span></span> |
| [<span data-ttu-id="24966-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="24966-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="24966-177">Método</span><span class="sxs-lookup"><span data-stu-id="24966-177">Method</span></span> |
| [<span data-ttu-id="24966-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="24966-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="24966-179">Método</span><span class="sxs-lookup"><span data-stu-id="24966-179">Method</span></span> |
| [<span data-ttu-id="24966-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="24966-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="24966-181">Método</span><span class="sxs-lookup"><span data-stu-id="24966-181">Method</span></span> |
| [<span data-ttu-id="24966-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="24966-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="24966-183">Método</span><span class="sxs-lookup"><span data-stu-id="24966-183">Method</span></span> |
| [<span data-ttu-id="24966-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="24966-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="24966-185">Método</span><span class="sxs-lookup"><span data-stu-id="24966-185">Method</span></span> |
| [<span data-ttu-id="24966-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="24966-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="24966-187">Método</span><span class="sxs-lookup"><span data-stu-id="24966-187">Method</span></span> |
| [<span data-ttu-id="24966-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="24966-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="24966-189">Método</span><span class="sxs-lookup"><span data-stu-id="24966-189">Method</span></span> |
| [<span data-ttu-id="24966-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="24966-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="24966-191">Método</span><span class="sxs-lookup"><span data-stu-id="24966-191">Method</span></span> |
| [<span data-ttu-id="24966-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="24966-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="24966-193">Método</span><span class="sxs-lookup"><span data-stu-id="24966-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="24966-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-194">Example</span></span>

<span data-ttu-id="24966-195">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="24966-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="24966-196">Members</span><span class="sxs-lookup"><span data-stu-id="24966-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="24966-197">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="24966-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="24966-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-200">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="24966-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="24966-201">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="24966-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="24966-202">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-202">Type</span></span>

*   <span data-ttu-id="24966-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="24966-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-204">Requirements</span></span>

|<span data-ttu-id="24966-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-205">Requirement</span></span>| <span data-ttu-id="24966-206">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-208">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-208">1.0</span></span>|
|[<span data-ttu-id="24966-209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-210">ReadItem</span></span>|
|[<span data-ttu-id="24966-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-212">Read</span><span class="sxs-lookup"><span data-stu-id="24966-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-213">Example</span></span>

<span data-ttu-id="24966-214">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="24966-215">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-215">bcc :[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-216">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="24966-217">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="24966-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-218">Type</span></span>

*   [<span data-ttu-id="24966-219">Destinatários</span><span class="sxs-lookup"><span data-stu-id="24966-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-220">Requirements</span></span>

|<span data-ttu-id="24966-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-221">Requirement</span></span>| <span data-ttu-id="24966-222">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-224">1.1</span><span class="sxs-lookup"><span data-stu-id="24966-224">1.1</span></span>|
|[<span data-ttu-id="24966-225">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-226">ReadItem</span></span>|
|[<span data-ttu-id="24966-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-228">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="24966-230">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-230">body :[Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-231">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="24966-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-232">Type</span></span>

*   [<span data-ttu-id="24966-233">Body</span><span class="sxs-lookup"><span data-stu-id="24966-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-234">Requirements</span></span>

|<span data-ttu-id="24966-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-235">Requirement</span></span>| <span data-ttu-id="24966-236">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-238">1.1</span><span class="sxs-lookup"><span data-stu-id="24966-238">1.1</span></span>|
|[<span data-ttu-id="24966-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-240">ReadItem</span></span>|
|[<span data-ttu-id="24966-241">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-243">Example</span></span>

<span data-ttu-id="24966-244">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="24966-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="24966-245">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="24966-246">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-247">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="24966-248">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-249">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-249">Read mode</span></span>

<span data-ttu-id="24966-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="24966-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-252">Compose mode</span></span>

<span data-ttu-id="24966-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="24966-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-254">Type</span></span>

*   <span data-ttu-id="24966-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-256">Requirements</span></span>

|<span data-ttu-id="24966-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-257">Requirement</span></span>| <span data-ttu-id="24966-258">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-260">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-260">1.0</span></span>|
|[<span data-ttu-id="24966-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-262">ReadItem</span></span>|
|[<span data-ttu-id="24966-263">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="24966-265">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="24966-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="24966-266">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="24966-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="24966-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="24966-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="24966-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="24966-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-271">Type</span></span>

*   <span data-ttu-id="24966-272">String</span><span class="sxs-lookup"><span data-stu-id="24966-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-273">Requirements</span></span>

|<span data-ttu-id="24966-274">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-274">Requirement</span></span>| <span data-ttu-id="24966-275">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-277">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-277">1.0</span></span>|
|[<span data-ttu-id="24966-278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-279">ReadItem</span></span>|
|[<span data-ttu-id="24966-280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-281">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="24966-283">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="24966-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="24966-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-286">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-286">Type</span></span>

*   <span data-ttu-id="24966-287">Data</span><span class="sxs-lookup"><span data-stu-id="24966-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-288">Requirements</span></span>

|<span data-ttu-id="24966-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-289">Requirement</span></span>| <span data-ttu-id="24966-290">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-292">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-292">1.0</span></span>|
|[<span data-ttu-id="24966-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-294">ReadItem</span></span>|
|[<span data-ttu-id="24966-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-296">Read</span><span class="sxs-lookup"><span data-stu-id="24966-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="24966-298">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="24966-298">dateTimeModified :Date</span></span>

<span data-ttu-id="24966-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-301">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-302">Type</span></span>

*   <span data-ttu-id="24966-303">Data</span><span class="sxs-lookup"><span data-stu-id="24966-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-304">Requirements</span></span>

|<span data-ttu-id="24966-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-305">Requirement</span></span>| <span data-ttu-id="24966-306">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-308">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-308">1.0</span></span>|
|[<span data-ttu-id="24966-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-310">ReadItem</span></span>|
|[<span data-ttu-id="24966-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-312">Read</span><span class="sxs-lookup"><span data-stu-id="24966-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="24966-314">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-314">end :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="24966-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="24966-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="24966-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-318">Read mode</span></span>

<span data-ttu-id="24966-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="24966-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-320">Compose mode</span></span>

<span data-ttu-id="24966-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="24966-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="24966-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="24966-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="24966-323">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="24966-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="24966-324">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-324">Type</span></span>

*   <span data-ttu-id="24966-325">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-326">Requirements</span></span>

|<span data-ttu-id="24966-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-327">Requirement</span></span>| <span data-ttu-id="24966-328">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-330">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-330">1.0</span></span>|
|[<span data-ttu-id="24966-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-332">ReadItem</span></span>|
|[<span data-ttu-id="24966-333">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="24966-335">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-335">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="24966-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="24966-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-340">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="24966-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-341">Type</span></span>

*   [<span data-ttu-id="24966-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="24966-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-343">Requirements</span></span>

|<span data-ttu-id="24966-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-344">Requirement</span></span>| <span data-ttu-id="24966-345">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-347">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-347">1.0</span></span>|
|[<span data-ttu-id="24966-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-349">ReadItem</span></span>|
|[<span data-ttu-id="24966-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-351">Read</span><span class="sxs-lookup"><span data-stu-id="24966-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="24966-353">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="24966-353">internetMessageId :String</span></span>

<span data-ttu-id="24966-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-356">Type</span></span>

*   <span data-ttu-id="24966-357">String</span><span class="sxs-lookup"><span data-stu-id="24966-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-358">Requirements</span></span>

|<span data-ttu-id="24966-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-359">Requirement</span></span>| <span data-ttu-id="24966-360">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-362">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-362">1.0</span></span>|
|[<span data-ttu-id="24966-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-364">ReadItem</span></span>|
|[<span data-ttu-id="24966-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-366">Read</span><span class="sxs-lookup"><span data-stu-id="24966-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="24966-368">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="24966-368">itemClass :String</span></span>

<span data-ttu-id="24966-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="24966-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="24966-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-373">Type</span></span> | <span data-ttu-id="24966-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-374">Description</span></span> | <span data-ttu-id="24966-375">classe de item</span><span class="sxs-lookup"><span data-stu-id="24966-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="24966-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="24966-376">Appointment items</span></span> | <span data-ttu-id="24966-377">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="24966-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="24966-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="24966-378">Message items</span></span> | <span data-ttu-id="24966-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="24966-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="24966-380">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="24966-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-381">Type</span></span>

*   <span data-ttu-id="24966-382">String</span><span class="sxs-lookup"><span data-stu-id="24966-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-383">Requirements</span></span>

|<span data-ttu-id="24966-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-384">Requirement</span></span>| <span data-ttu-id="24966-385">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-387">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-387">1.0</span></span>|
|[<span data-ttu-id="24966-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-389">ReadItem</span></span>|
|[<span data-ttu-id="24966-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-391">Read</span><span class="sxs-lookup"><span data-stu-id="24966-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="24966-393">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24966-393">(nullable) itemId :String</span></span>

<span data-ttu-id="24966-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="24966-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="24966-397">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="24966-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="24966-398">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="24966-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="24966-399">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="24966-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="24966-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-402">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-402">Type</span></span>

*   <span data-ttu-id="24966-403">String</span><span class="sxs-lookup"><span data-stu-id="24966-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-404">Requirements</span></span>

|<span data-ttu-id="24966-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-405">Requirement</span></span>| <span data-ttu-id="24966-406">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-407">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-408">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-408">1.0</span></span>|
|[<span data-ttu-id="24966-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-410">ReadItem</span></span>|
|[<span data-ttu-id="24966-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-412">Read</span><span class="sxs-lookup"><span data-stu-id="24966-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-413">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-413">Example</span></span>

<span data-ttu-id="24966-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="24966-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="24966-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-417">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="24966-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="24966-418">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-419">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-419">Type</span></span>

*   [<span data-ttu-id="24966-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="24966-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-421">Requirements</span></span>

|<span data-ttu-id="24966-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-422">Requirement</span></span>| <span data-ttu-id="24966-423">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-425">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-425">1.0</span></span>|
|[<span data-ttu-id="24966-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-427">ReadItem</span></span>|
|[<span data-ttu-id="24966-428">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-429">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="24966-431">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-431">location :String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-432">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-433">Read mode</span></span>

<span data-ttu-id="24966-434">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-435">Compose mode</span></span>

<span data-ttu-id="24966-436">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="24966-437">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-437">Type</span></span>

*   <span data-ttu-id="24966-438">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-439">Requirements</span></span>

|<span data-ttu-id="24966-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-440">Requirement</span></span>| <span data-ttu-id="24966-441">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-443">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-443">1.0</span></span>|
|[<span data-ttu-id="24966-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-445">ReadItem</span></span>|
|[<span data-ttu-id="24966-446">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-447">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="24966-448">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24966-448">normalizedSubject :String</span></span>

<span data-ttu-id="24966-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="24966-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="24966-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-453">Type</span></span>

*   <span data-ttu-id="24966-454">String</span><span class="sxs-lookup"><span data-stu-id="24966-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-455">Requirements</span></span>

|<span data-ttu-id="24966-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-456">Requirement</span></span>| <span data-ttu-id="24966-457">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-459">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-459">1.0</span></span>|
|[<span data-ttu-id="24966-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-461">ReadItem</span></span>|
|[<span data-ttu-id="24966-462">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-463">Read</span><span class="sxs-lookup"><span data-stu-id="24966-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-464">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="24966-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-465">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-466">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="24966-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-467">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-467">Type</span></span>

*   [<span data-ttu-id="24966-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="24966-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-469">Requirements</span></span>

|<span data-ttu-id="24966-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-470">Requirement</span></span>| <span data-ttu-id="24966-471">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-473">1.3</span><span class="sxs-lookup"><span data-stu-id="24966-473">1.3</span></span>|
|[<span data-ttu-id="24966-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-475">ReadItem</span></span>|
|[<span data-ttu-id="24966-476">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-477">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="24966-479">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-480">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="24966-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="24966-481">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-482">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-482">Read mode</span></span>

<span data-ttu-id="24966-483">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="24966-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-484">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-484">Compose mode</span></span>

<span data-ttu-id="24966-485">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="24966-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="24966-486">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-486">Type</span></span>

*   <span data-ttu-id="24966-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-488">Requirements</span></span>

|<span data-ttu-id="24966-489">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-489">Requirement</span></span>| <span data-ttu-id="24966-490">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-491">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-492">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-492">1.0</span></span>|
|[<span data-ttu-id="24966-493">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-494">ReadItem</span></span>|
|[<span data-ttu-id="24966-495">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-496">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="24966-497">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-497">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-500">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-500">Type</span></span>

*   [<span data-ttu-id="24966-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="24966-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-502">Requirements</span></span>

|<span data-ttu-id="24966-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-503">Requirement</span></span>| <span data-ttu-id="24966-504">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-506">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-506">1.0</span></span>|
|[<span data-ttu-id="24966-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-508">ReadItem</span></span>|
|[<span data-ttu-id="24966-509">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-510">Read</span><span class="sxs-lookup"><span data-stu-id="24966-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-511">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="24966-512">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-513">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="24966-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="24966-514">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-515">Read mode</span></span>

<span data-ttu-id="24966-516">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="24966-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-517">Compose mode</span></span>

<span data-ttu-id="24966-518">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="24966-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="24966-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-519">Type</span></span>

*   <span data-ttu-id="24966-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-521">Requirements</span></span>

|<span data-ttu-id="24966-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-522">Requirement</span></span>| <span data-ttu-id="24966-523">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-525">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-525">1.0</span></span>|
|[<span data-ttu-id="24966-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-527">ReadItem</span></span>|
|[<span data-ttu-id="24966-528">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-529">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="24966-530">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-530">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="24966-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="24966-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="24966-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-535">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="24966-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="24966-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-536">Type</span></span>

*   [<span data-ttu-id="24966-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="24966-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="24966-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-538">Requirements</span></span>

|<span data-ttu-id="24966-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-539">Requirement</span></span>| <span data-ttu-id="24966-540">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-542">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-542">1.0</span></span>|
|[<span data-ttu-id="24966-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-544">ReadItem</span></span>|
|[<span data-ttu-id="24966-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-546">Read</span><span class="sxs-lookup"><span data-stu-id="24966-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="24966-548">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-548">start :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-549">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="24966-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="24966-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="24966-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-552">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-552">Read mode</span></span>

<span data-ttu-id="24966-553">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="24966-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-554">Compose mode</span></span>

<span data-ttu-id="24966-555">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="24966-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="24966-556">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="24966-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="24966-557">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="24966-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="24966-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-558">Type</span></span>

*   <span data-ttu-id="24966-559">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-560">Requirements</span></span>

|<span data-ttu-id="24966-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-561">Requirement</span></span>| <span data-ttu-id="24966-562">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-564">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-564">1.0</span></span>|
|[<span data-ttu-id="24966-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-566">ReadItem</span></span>|
|[<span data-ttu-id="24966-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="24966-569">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-569">subject :String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="24966-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="24966-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="24966-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-572">Read mode</span></span>

<span data-ttu-id="24966-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="24966-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-575">Compose mode</span></span>

<span data-ttu-id="24966-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="24966-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="24966-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-577">Type</span></span>

*   <span data-ttu-id="24966-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-579">Requirements</span></span>

|<span data-ttu-id="24966-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-580">Requirement</span></span>| <span data-ttu-id="24966-581">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-583">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-583">1.0</span></span>|
|[<span data-ttu-id="24966-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-585">ReadItem</span></span>|
|[<span data-ttu-id="24966-586">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-587">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="24966-588">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="24966-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="24966-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="24966-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="24966-591">Read mode</span></span>

<span data-ttu-id="24966-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="24966-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="24966-594">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="24966-594">Compose mode</span></span>

<span data-ttu-id="24966-595">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="24966-596">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-596">Type</span></span>

*   <span data-ttu-id="24966-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-598">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-598">Requirements</span></span>

|<span data-ttu-id="24966-599">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-599">Requirement</span></span>| <span data-ttu-id="24966-600">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-601">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-602">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-602">1.0</span></span>|
|[<span data-ttu-id="24966-603">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-604">ReadItem</span></span>|
|[<span data-ttu-id="24966-605">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-606">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="24966-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="24966-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="24966-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="24966-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="24966-609">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="24966-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="24966-610">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="24966-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="24966-611">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="24966-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-612">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-612">Parameters</span></span>

|<span data-ttu-id="24966-613">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-613">Name</span></span>| <span data-ttu-id="24966-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-614">Type</span></span>| <span data-ttu-id="24966-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-615">Attributes</span></span>| <span data-ttu-id="24966-616">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="24966-617">String</span><span class="sxs-lookup"><span data-stu-id="24966-617">String</span></span>||<span data-ttu-id="24966-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="24966-620">String</span><span class="sxs-lookup"><span data-stu-id="24966-620">String</span></span>||<span data-ttu-id="24966-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="24966-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-623">Object</span></span>| <span data-ttu-id="24966-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-624">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="24966-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-626">Object</span></span> | <span data-ttu-id="24966-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-627">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="24966-629">Booliano</span><span class="sxs-lookup"><span data-stu-id="24966-629">Boolean</span></span> | <span data-ttu-id="24966-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-630">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-631">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="24966-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="24966-632">function</span><span class="sxs-lookup"><span data-stu-id="24966-632">function</span></span>| <span data-ttu-id="24966-633">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-633">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-634">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="24966-635">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="24966-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="24966-636">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="24966-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="24966-637">Erros</span><span class="sxs-lookup"><span data-stu-id="24966-637">Errors</span></span>

| <span data-ttu-id="24966-638">Código de erro</span><span class="sxs-lookup"><span data-stu-id="24966-638">Error code</span></span> | <span data-ttu-id="24966-639">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="24966-640">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="24966-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="24966-641">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="24966-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="24966-642">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="24966-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-643">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-643">Requirements</span></span>

|<span data-ttu-id="24966-644">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-644">Requirement</span></span>| <span data-ttu-id="24966-645">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-646">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-647">1.1</span><span class="sxs-lookup"><span data-stu-id="24966-647">1.1</span></span>|
|[<span data-ttu-id="24966-648">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-650">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-651">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="24966-652">Exemplos</span><span class="sxs-lookup"><span data-stu-id="24966-652">Examples</span></span>

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

<span data-ttu-id="24966-653">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="24966-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="24966-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="24966-655">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="24966-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="24966-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="24966-659">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="24966-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="24966-660">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="24966-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-661">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-661">Parameters</span></span>

|<span data-ttu-id="24966-662">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-662">Name</span></span>| <span data-ttu-id="24966-663">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-663">Type</span></span>| <span data-ttu-id="24966-664">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-664">Attributes</span></span>| <span data-ttu-id="24966-665">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="24966-666">String</span><span class="sxs-lookup"><span data-stu-id="24966-666">String</span></span>||<span data-ttu-id="24966-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="24966-669">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24966-669">String</span></span>||<span data-ttu-id="24966-670">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="24966-670">The subject of the item to be attached.</span></span> <span data-ttu-id="24966-671">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="24966-672">Object</span><span class="sxs-lookup"><span data-stu-id="24966-672">Object</span></span>| <span data-ttu-id="24966-673">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-673">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-674">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="24966-675">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-675">Object</span></span>| <span data-ttu-id="24966-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-676">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-677">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="24966-678">function</span><span class="sxs-lookup"><span data-stu-id="24966-678">function</span></span>| <span data-ttu-id="24966-679">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-679">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-680">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="24966-681">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="24966-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="24966-682">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="24966-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="24966-683">Erros</span><span class="sxs-lookup"><span data-stu-id="24966-683">Errors</span></span>

| <span data-ttu-id="24966-684">Código de erro</span><span class="sxs-lookup"><span data-stu-id="24966-684">Error code</span></span> | <span data-ttu-id="24966-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="24966-686">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="24966-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-687">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-687">Requirements</span></span>

|<span data-ttu-id="24966-688">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-688">Requirement</span></span>| <span data-ttu-id="24966-689">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-690">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-691">1.1</span><span class="sxs-lookup"><span data-stu-id="24966-691">1.1</span></span>|
|[<span data-ttu-id="24966-692">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-694">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-695">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-696">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-696">Example</span></span>

<span data-ttu-id="24966-697">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="24966-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="24966-698">close()</span><span class="sxs-lookup"><span data-stu-id="24966-698">close()</span></span>

<span data-ttu-id="24966-699">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="24966-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="24966-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="24966-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-702">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="24966-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="24966-703">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="24966-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-704">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-704">Requirements</span></span>

|<span data-ttu-id="24966-705">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-705">Requirement</span></span>| <span data-ttu-id="24966-706">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-707">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-708">1.3</span><span class="sxs-lookup"><span data-stu-id="24966-708">1.3</span></span>|
|[<span data-ttu-id="24966-709">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-710">Restrito</span><span class="sxs-lookup"><span data-stu-id="24966-710">Restricted</span></span>|
|[<span data-ttu-id="24966-711">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-712">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="24966-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="24966-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="24966-714">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="24966-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-715">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="24966-716">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="24966-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="24966-717">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="24966-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="24966-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="24966-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-721">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-721">Parameters</span></span>

| <span data-ttu-id="24966-722">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-722">Name</span></span> | <span data-ttu-id="24966-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-723">Type</span></span> | <span data-ttu-id="24966-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-724">Attributes</span></span> | <span data-ttu-id="24966-725">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="24966-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="24966-726">String &#124; Object</span></span>| |<span data-ttu-id="24966-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="24966-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="24966-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="24966-729">**OR**</span></span><br/><span data-ttu-id="24966-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="24966-732">String</span><span class="sxs-lookup"><span data-stu-id="24966-732">String</span></span> | <span data-ttu-id="24966-733">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-733">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="24966-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="24966-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="24966-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-737">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-738">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="24966-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="24966-739">String</span><span class="sxs-lookup"><span data-stu-id="24966-739">String</span></span> | | <span data-ttu-id="24966-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="24966-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="24966-742">String</span><span class="sxs-lookup"><span data-stu-id="24966-742">String</span></span> | | <span data-ttu-id="24966-743">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="24966-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="24966-744">String</span><span class="sxs-lookup"><span data-stu-id="24966-744">String</span></span> | | <span data-ttu-id="24966-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="24966-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="24966-747">Booliano</span><span class="sxs-lookup"><span data-stu-id="24966-747">Boolean</span></span> | | <span data-ttu-id="24966-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="24966-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="24966-750">String</span><span class="sxs-lookup"><span data-stu-id="24966-750">String</span></span> | | <span data-ttu-id="24966-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="24966-754">function</span><span class="sxs-lookup"><span data-stu-id="24966-754">function</span></span> | <span data-ttu-id="24966-755">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-755">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-756">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-757">Requirements</span></span>

|<span data-ttu-id="24966-758">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-758">Requirement</span></span>| <span data-ttu-id="24966-759">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-760">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-761">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-761">1.0</span></span>|
|[<span data-ttu-id="24966-762">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-763">ReadItem</span></span>|
|[<span data-ttu-id="24966-764">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-765">Read</span><span class="sxs-lookup"><span data-stu-id="24966-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="24966-766">Exemplos</span><span class="sxs-lookup"><span data-stu-id="24966-766">Examples</span></span>

<span data-ttu-id="24966-767">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="24966-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="24966-768">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="24966-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="24966-769">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="24966-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="24966-770">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="24966-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="24966-771">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="24966-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="24966-772">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="24966-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="24966-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="24966-774">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="24966-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-775">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="24966-776">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="24966-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="24966-777">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="24966-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="24966-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="24966-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-781">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-781">Parameters</span></span>

| <span data-ttu-id="24966-782">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-782">Name</span></span> | <span data-ttu-id="24966-783">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-783">Type</span></span> | <span data-ttu-id="24966-784">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-784">Attributes</span></span> | <span data-ttu-id="24966-785">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="24966-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="24966-786">String &#124; Object</span></span>| | <span data-ttu-id="24966-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="24966-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="24966-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="24966-789">**OR**</span></span><br/><span data-ttu-id="24966-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="24966-792">String</span><span class="sxs-lookup"><span data-stu-id="24966-792">String</span></span> | <span data-ttu-id="24966-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-793">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="24966-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="24966-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="24966-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-797">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-798">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="24966-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="24966-799">String</span><span class="sxs-lookup"><span data-stu-id="24966-799">String</span></span> | | <span data-ttu-id="24966-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="24966-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="24966-802">String</span><span class="sxs-lookup"><span data-stu-id="24966-802">String</span></span> | | <span data-ttu-id="24966-803">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="24966-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="24966-804">String</span><span class="sxs-lookup"><span data-stu-id="24966-804">String</span></span> | | <span data-ttu-id="24966-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="24966-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="24966-807">Booliano</span><span class="sxs-lookup"><span data-stu-id="24966-807">Boolean</span></span> | | <span data-ttu-id="24966-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="24966-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="24966-810">String</span><span class="sxs-lookup"><span data-stu-id="24966-810">String</span></span> | | <span data-ttu-id="24966-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="24966-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="24966-814">function</span><span class="sxs-lookup"><span data-stu-id="24966-814">function</span></span> | <span data-ttu-id="24966-815">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-815">&lt;optional&gt;</span></span> | <span data-ttu-id="24966-816">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-817">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-817">Requirements</span></span>

|<span data-ttu-id="24966-818">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-818">Requirement</span></span>| <span data-ttu-id="24966-819">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-820">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-821">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-821">1.0</span></span>|
|[<span data-ttu-id="24966-822">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-823">ReadItem</span></span>|
|[<span data-ttu-id="24966-824">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-825">Read</span><span class="sxs-lookup"><span data-stu-id="24966-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="24966-826">Exemplos</span><span class="sxs-lookup"><span data-stu-id="24966-826">Examples</span></span>

<span data-ttu-id="24966-827">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="24966-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="24966-828">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="24966-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="24966-829">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="24966-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="24966-830">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="24966-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="24966-831">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="24966-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="24966-832">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="24966-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="24966-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="24966-834">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="24966-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-835">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-836">Requirements</span></span>

|<span data-ttu-id="24966-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-837">Requirement</span></span>| <span data-ttu-id="24966-838">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-840">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-840">1.0</span></span>|
|[<span data-ttu-id="24966-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-842">ReadItem</span></span>|
|[<span data-ttu-id="24966-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-844">Read</span><span class="sxs-lookup"><span data-stu-id="24966-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-845">Returns:</span></span>

<span data-ttu-id="24966-846">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="24966-846">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="24966-847">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-847">Example</span></span>

<span data-ttu-id="24966-848">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="24966-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="24966-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="24966-850">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="24966-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-851">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-852">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-852">Parameters</span></span>

|<span data-ttu-id="24966-853">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-853">Name</span></span>| <span data-ttu-id="24966-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-854">Type</span></span>| <span data-ttu-id="24966-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="24966-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="24966-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="24966-857">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="24966-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-858">Requirements</span></span>

|<span data-ttu-id="24966-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-859">Requirement</span></span>| <span data-ttu-id="24966-860">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-862">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-862">1.0</span></span>|
|[<span data-ttu-id="24966-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-864">Restrito</span><span class="sxs-lookup"><span data-stu-id="24966-864">Restricted</span></span>|
|[<span data-ttu-id="24966-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-866">Read</span><span class="sxs-lookup"><span data-stu-id="24966-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-867">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-867">Returns:</span></span>

<span data-ttu-id="24966-868">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="24966-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="24966-869">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="24966-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="24966-870">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="24966-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="24966-871">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="24966-872">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="24966-872">Value of `entityType`</span></span> | <span data-ttu-id="24966-873">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="24966-873">Type of objects in returned array</span></span> | <span data-ttu-id="24966-874">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="24966-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="24966-875">String</span><span class="sxs-lookup"><span data-stu-id="24966-875">String</span></span> | <span data-ttu-id="24966-876">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="24966-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="24966-877">Contato</span><span class="sxs-lookup"><span data-stu-id="24966-877">Contact</span></span> | <span data-ttu-id="24966-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="24966-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="24966-879">String</span><span class="sxs-lookup"><span data-stu-id="24966-879">String</span></span> | <span data-ttu-id="24966-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="24966-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="24966-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="24966-881">MeetingSuggestion</span></span> | <span data-ttu-id="24966-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="24966-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="24966-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="24966-883">PhoneNumber</span></span> | <span data-ttu-id="24966-884">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="24966-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="24966-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="24966-885">TaskSuggestion</span></span> | <span data-ttu-id="24966-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="24966-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="24966-887">String</span><span class="sxs-lookup"><span data-stu-id="24966-887">String</span></span> | <span data-ttu-id="24966-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="24966-888">**Restricted**</span></span> |

<span data-ttu-id="24966-889">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="24966-889">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="24966-890">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-890">Example</span></span>

<span data-ttu-id="24966-891">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="24966-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficeviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="24966-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="24966-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="24966-893">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="24966-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-894">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="24966-895">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="24966-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-896">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-896">Parameters</span></span>

|<span data-ttu-id="24966-897">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-897">Name</span></span>| <span data-ttu-id="24966-898">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-898">Type</span></span>| <span data-ttu-id="24966-899">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="24966-900">String</span><span class="sxs-lookup"><span data-stu-id="24966-900">String</span></span>|<span data-ttu-id="24966-901">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="24966-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-902">Requirements</span></span>

|<span data-ttu-id="24966-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-903">Requirement</span></span>| <span data-ttu-id="24966-904">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-906">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-906">1.0</span></span>|
|[<span data-ttu-id="24966-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-908">ReadItem</span></span>|
|[<span data-ttu-id="24966-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-910">Read</span><span class="sxs-lookup"><span data-stu-id="24966-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-911">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-911">Returns:</span></span>

<span data-ttu-id="24966-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="24966-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="24966-914">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="24966-914">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="24966-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="24966-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="24966-916">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="24966-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-917">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="24966-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="24966-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="24966-921">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="24966-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="24966-922">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="24966-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="24966-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="24966-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="24966-926">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-926">Requirements</span></span>

|<span data-ttu-id="24966-927">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-927">Requirement</span></span>| <span data-ttu-id="24966-928">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-929">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-930">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-930">1.0</span></span>|
|[<span data-ttu-id="24966-931">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-932">ReadItem</span></span>|
|[<span data-ttu-id="24966-933">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-934">Read</span><span class="sxs-lookup"><span data-stu-id="24966-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-935">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-935">Returns:</span></span>

<span data-ttu-id="24966-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="24966-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="24966-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="24966-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="24966-939">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="24966-940">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-940">Example</span></span>

<span data-ttu-id="24966-941">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="24966-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="24966-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="24966-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="24966-943">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="24966-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-944">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="24966-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="24966-945">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="24966-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="24966-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="24966-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-948">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-948">Parameters</span></span>

|<span data-ttu-id="24966-949">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-949">Name</span></span>| <span data-ttu-id="24966-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-950">Type</span></span>| <span data-ttu-id="24966-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="24966-952">String</span><span class="sxs-lookup"><span data-stu-id="24966-952">String</span></span>|<span data-ttu-id="24966-953">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="24966-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-954">Requirements</span></span>

|<span data-ttu-id="24966-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-955">Requirement</span></span>| <span data-ttu-id="24966-956">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-957">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-958">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-958">1.0</span></span>|
|[<span data-ttu-id="24966-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-959">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-960">ReadItem</span></span>|
|[<span data-ttu-id="24966-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-961">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-962">Read</span><span class="sxs-lookup"><span data-stu-id="24966-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-963">Returns:</span></span>

<span data-ttu-id="24966-964">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="24966-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="24966-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="24966-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="24966-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="24966-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="24966-967">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="24966-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="24966-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="24966-969">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="24966-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="24966-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-972">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-972">Parameters</span></span>

|<span data-ttu-id="24966-973">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-973">Name</span></span>| <span data-ttu-id="24966-974">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-974">Type</span></span>| <span data-ttu-id="24966-975">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-975">Attributes</span></span>| <span data-ttu-id="24966-976">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="24966-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="24966-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="24966-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="24966-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="24966-981">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-981">Object</span></span>| <span data-ttu-id="24966-982">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-982">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-983">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="24966-984">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-984">Object</span></span>| <span data-ttu-id="24966-985">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-985">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-986">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="24966-987">function</span><span class="sxs-lookup"><span data-stu-id="24966-987">function</span></span>||<span data-ttu-id="24966-988">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="24966-989">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="24966-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="24966-990">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="24966-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-991">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-991">Requirements</span></span>

|<span data-ttu-id="24966-992">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-992">Requirement</span></span>| <span data-ttu-id="24966-993">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-994">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-995">1.2</span><span class="sxs-lookup"><span data-stu-id="24966-995">1.2</span></span>|
|[<span data-ttu-id="24966-996">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-996">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-998">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-998">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-999">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="24966-1000">Retorna:</span><span class="sxs-lookup"><span data-stu-id="24966-1000">Returns:</span></span>

<span data-ttu-id="24966-1001">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="24966-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="24966-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="24966-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="24966-1003">String</span><span class="sxs-lookup"><span data-stu-id="24966-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="24966-1004">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-1004">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="24966-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="24966-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="24966-1006">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="24966-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="24966-p163">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="24966-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-1010">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-1010">Parameters</span></span>

|<span data-ttu-id="24966-1011">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-1011">Name</span></span>| <span data-ttu-id="24966-1012">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-1012">Type</span></span>| <span data-ttu-id="24966-1013">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-1013">Attributes</span></span>| <span data-ttu-id="24966-1014">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="24966-1015">function</span><span class="sxs-lookup"><span data-stu-id="24966-1015">function</span></span>||<span data-ttu-id="24966-1016">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="24966-1017">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="24966-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="24966-1018">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="24966-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="24966-1019">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1019">Object</span></span>| <span data-ttu-id="24966-1020">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1021">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="24966-1022">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-1023">Requirements</span></span>

|<span data-ttu-id="24966-1024">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-1024">Requirement</span></span>| <span data-ttu-id="24966-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-1026">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="24966-1027">1.0</span></span>|
|[<span data-ttu-id="24966-1028">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-1028">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24966-1029">ReadItem</span></span>|
|[<span data-ttu-id="24966-1030">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="24966-1030">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-1031">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24966-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-1032">Example</span></span>

<span data-ttu-id="24966-p166">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="24966-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="24966-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="24966-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="24966-1037">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="24966-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="24966-p167">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook na Web e em dispositivos móveis, o identificador do anexo é valido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, posteriormente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="24966-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-1042">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-1042">Parameters</span></span>

|<span data-ttu-id="24966-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-1043">Name</span></span>| <span data-ttu-id="24966-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-1044">Type</span></span>| <span data-ttu-id="24966-1045">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-1045">Attributes</span></span>| <span data-ttu-id="24966-1046">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="24966-1047">String</span><span class="sxs-lookup"><span data-stu-id="24966-1047">String</span></span>||<span data-ttu-id="24966-1048">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="24966-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="24966-1049">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1049">Object</span></span>| <span data-ttu-id="24966-1050">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1051">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="24966-1052">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1052">Object</span></span>| <span data-ttu-id="24966-1053">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1054">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="24966-1055">function</span><span class="sxs-lookup"><span data-stu-id="24966-1055">function</span></span>| <span data-ttu-id="24966-1056">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1057">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="24966-1058">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="24966-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="24966-1059">Erros</span><span class="sxs-lookup"><span data-stu-id="24966-1059">Errors</span></span>

| <span data-ttu-id="24966-1060">Código de erro</span><span class="sxs-lookup"><span data-stu-id="24966-1060">Error code</span></span> | <span data-ttu-id="24966-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="24966-1062">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="24966-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-1063">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-1063">Requirements</span></span>

|<span data-ttu-id="24966-1064">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-1064">Requirement</span></span>| <span data-ttu-id="24966-1065">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-1066">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="24966-1067">1.1</span></span>|
|[<span data-ttu-id="24966-1068">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-1070">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-1071">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-1072">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-1072">Example</span></span>

<span data-ttu-id="24966-1073">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="24966-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="24966-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="24966-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="24966-1075">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="24966-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="24966-p168">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook na Web ou Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="24966-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-1079">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="24966-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="24966-1080">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="24966-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="24966-p170">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="24966-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="24966-1084">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="24966-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="24966-1085">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="24966-1085">Note: Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="24966-1086">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="24966-1086">The `saveAsync` method will fail when called from a meeting in compose mode.</span></span> <span data-ttu-id="24966-1087">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="24966-1087">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="24966-1088">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="24966-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-1089">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-1089">Parameters</span></span>

|<span data-ttu-id="24966-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-1090">Name</span></span>| <span data-ttu-id="24966-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-1091">Type</span></span>| <span data-ttu-id="24966-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-1092">Attributes</span></span>| <span data-ttu-id="24966-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="24966-1094">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1094">Object</span></span>| <span data-ttu-id="24966-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1096">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="24966-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1097">Object</span></span>| <span data-ttu-id="24966-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1099">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="24966-1100">function</span><span class="sxs-lookup"><span data-stu-id="24966-1100">function</span></span>||<span data-ttu-id="24966-1101">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="24966-1102">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="24966-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24966-1103">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-1103">Requirements</span></span>

|<span data-ttu-id="24966-1104">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-1104">Requirement</span></span>| <span data-ttu-id="24966-1105">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-1106">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="24966-1107">1.3</span></span>|
|[<span data-ttu-id="24966-1108">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-1108">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-1110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-1110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-1111">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="24966-1112">Exemplos</span><span class="sxs-lookup"><span data-stu-id="24966-1112">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="24966-p172">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="24966-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="24966-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="24966-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="24966-1116">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="24966-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="24966-p173">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="24966-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="24966-1120">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="24966-1120">Parameters</span></span>

|<span data-ttu-id="24966-1121">Nome</span><span class="sxs-lookup"><span data-stu-id="24966-1121">Name</span></span>| <span data-ttu-id="24966-1122">Tipo</span><span class="sxs-lookup"><span data-stu-id="24966-1122">Type</span></span>| <span data-ttu-id="24966-1123">Atributos</span><span class="sxs-lookup"><span data-stu-id="24966-1123">Attributes</span></span>| <span data-ttu-id="24966-1124">Descrição</span><span class="sxs-lookup"><span data-stu-id="24966-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="24966-1125">String</span><span class="sxs-lookup"><span data-stu-id="24966-1125">String</span></span>||<span data-ttu-id="24966-p174">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="24966-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="24966-1129">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1129">Object</span></span>| <span data-ttu-id="24966-1130">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1131">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="24966-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="24966-1132">Objeto</span><span class="sxs-lookup"><span data-stu-id="24966-1132">Object</span></span>| <span data-ttu-id="24966-1133">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1134">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="24966-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="24966-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="24966-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="24966-1136">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="24966-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="24966-1137">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="24966-1137">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="24966-1138">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="24966-1138">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="24966-1139">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="24966-1139">If `html` and the field supports HTML (the subject doesn&#39;t), the current style is applied in Outlook Web App and the default style is applied in Outlook.</span></span> <span data-ttu-id="24966-1140">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="24966-1140">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="24966-1141">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="24966-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="24966-1142">function</span><span class="sxs-lookup"><span data-stu-id="24966-1142">function</span></span>||<span data-ttu-id="24966-1143">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="24966-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="24966-1144">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24966-1144">Requirements</span></span>

|<span data-ttu-id="24966-1145">Requisito</span><span class="sxs-lookup"><span data-stu-id="24966-1145">Requirement</span></span>| <span data-ttu-id="24966-1146">Valor</span><span class="sxs-lookup"><span data-stu-id="24966-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="24966-1147">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24966-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24966-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="24966-1148">1.2</span></span>|
|[<span data-ttu-id="24966-1149">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24966-1149">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24966-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="24966-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="24966-1151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24966-1151">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24966-1152">Escrever</span><span class="sxs-lookup"><span data-stu-id="24966-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="24966-1153">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24966-1153">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
