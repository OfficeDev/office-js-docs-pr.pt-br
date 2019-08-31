---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: dfea6e86372c2a4310b5fa458ff3dcede7c6a184
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696278"
---
# <a name="item"></a><span data-ttu-id="2141d-102">item</span><span class="sxs-lookup"><span data-stu-id="2141d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2141d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2141d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2141d-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="2141d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-106">Requirements</span></span>

|<span data-ttu-id="2141d-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-107">Requirement</span></span>| <span data-ttu-id="2141d-108">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-110">1.0</span></span>|
|[<span data-ttu-id="2141d-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="2141d-112">Restricted</span></span>|
|[<span data-ttu-id="2141d-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2141d-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="2141d-115">Members and methods</span></span>

| <span data-ttu-id="2141d-116">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-116">Member</span></span> | <span data-ttu-id="2141d-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2141d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="2141d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2141d-119">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-119">Member</span></span> |
| [<span data-ttu-id="2141d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="2141d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2141d-121">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-121">Member</span></span> |
| [<span data-ttu-id="2141d-122">body</span><span class="sxs-lookup"><span data-stu-id="2141d-122">body</span></span>](#body-body) | <span data-ttu-id="2141d-123">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-123">Member</span></span> |
| [<span data-ttu-id="2141d-124">cc</span><span class="sxs-lookup"><span data-stu-id="2141d-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2141d-125">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-125">Member</span></span> |
| [<span data-ttu-id="2141d-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="2141d-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2141d-127">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-127">Member</span></span> |
| [<span data-ttu-id="2141d-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2141d-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2141d-129">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-129">Member</span></span> |
| [<span data-ttu-id="2141d-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2141d-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2141d-131">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-131">Member</span></span> |
| [<span data-ttu-id="2141d-132">end</span><span class="sxs-lookup"><span data-stu-id="2141d-132">end</span></span>](#end-datetime) | <span data-ttu-id="2141d-133">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-133">Member</span></span> |
| [<span data-ttu-id="2141d-134">from</span><span class="sxs-lookup"><span data-stu-id="2141d-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="2141d-135">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-135">Member</span></span> |
| [<span data-ttu-id="2141d-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2141d-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2141d-137">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-137">Member</span></span> |
| [<span data-ttu-id="2141d-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="2141d-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2141d-139">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-139">Member</span></span> |
| [<span data-ttu-id="2141d-140">itemId</span><span class="sxs-lookup"><span data-stu-id="2141d-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2141d-141">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-141">Member</span></span> |
| [<span data-ttu-id="2141d-142">itemType</span><span class="sxs-lookup"><span data-stu-id="2141d-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2141d-143">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-143">Member</span></span> |
| [<span data-ttu-id="2141d-144">location</span><span class="sxs-lookup"><span data-stu-id="2141d-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="2141d-145">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-145">Member</span></span> |
| [<span data-ttu-id="2141d-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2141d-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2141d-147">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-147">Member</span></span> |
| [<span data-ttu-id="2141d-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2141d-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="2141d-149">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-149">Member</span></span> |
| [<span data-ttu-id="2141d-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2141d-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2141d-151">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-151">Member</span></span> |
| [<span data-ttu-id="2141d-152">organizer</span><span class="sxs-lookup"><span data-stu-id="2141d-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="2141d-153">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-153">Member</span></span> |
| [<span data-ttu-id="2141d-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2141d-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2141d-155">Member</span><span class="sxs-lookup"><span data-stu-id="2141d-155">Member</span></span> |
| [<span data-ttu-id="2141d-156">sender</span><span class="sxs-lookup"><span data-stu-id="2141d-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2141d-157">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-157">Member</span></span> |
| [<span data-ttu-id="2141d-158">start</span><span class="sxs-lookup"><span data-stu-id="2141d-158">start</span></span>](#start-datetime) | <span data-ttu-id="2141d-159">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-159">Member</span></span> |
| [<span data-ttu-id="2141d-160">subject</span><span class="sxs-lookup"><span data-stu-id="2141d-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2141d-161">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-161">Member</span></span> |
| [<span data-ttu-id="2141d-162">to</span><span class="sxs-lookup"><span data-stu-id="2141d-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2141d-163">Membro</span><span class="sxs-lookup"><span data-stu-id="2141d-163">Member</span></span> |
| [<span data-ttu-id="2141d-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2141d-165">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-165">Method</span></span> |
| [<span data-ttu-id="2141d-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2141d-167">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-167">Method</span></span> |
| [<span data-ttu-id="2141d-168">close</span><span class="sxs-lookup"><span data-stu-id="2141d-168">close</span></span>](#close) | <span data-ttu-id="2141d-169">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-169">Method</span></span> |
| [<span data-ttu-id="2141d-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2141d-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2141d-171">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-171">Method</span></span> |
| [<span data-ttu-id="2141d-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2141d-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2141d-173">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-173">Method</span></span> |
| [<span data-ttu-id="2141d-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="2141d-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2141d-175">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-175">Method</span></span> |
| [<span data-ttu-id="2141d-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2141d-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2141d-177">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-177">Method</span></span> |
| [<span data-ttu-id="2141d-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2141d-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2141d-179">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-179">Method</span></span> |
| [<span data-ttu-id="2141d-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2141d-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2141d-181">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-181">Method</span></span> |
| [<span data-ttu-id="2141d-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2141d-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2141d-183">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-183">Method</span></span> |
| [<span data-ttu-id="2141d-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2141d-185">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-185">Method</span></span> |
| [<span data-ttu-id="2141d-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2141d-187">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-187">Method</span></span> |
| [<span data-ttu-id="2141d-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2141d-189">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-189">Method</span></span> |
| [<span data-ttu-id="2141d-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2141d-191">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-191">Method</span></span> |
| [<span data-ttu-id="2141d-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2141d-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2141d-193">Método</span><span class="sxs-lookup"><span data-stu-id="2141d-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2141d-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-194">Example</span></span>

<span data-ttu-id="2141d-195">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="2141d-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2141d-196">Membros</span><span class="sxs-lookup"><span data-stu-id="2141d-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="2141d-197">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="2141d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="2141d-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-200">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="2141d-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2141d-201">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="2141d-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-202">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-202">Type</span></span>

*   <span data-ttu-id="2141d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="2141d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-204">Requirements</span></span>

|<span data-ttu-id="2141d-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-205">Requirement</span></span>| <span data-ttu-id="2141d-206">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-208">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-208">1.0</span></span>|
|[<span data-ttu-id="2141d-209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-210">ReadItem</span></span>|
|[<span data-ttu-id="2141d-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-212">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-213">Example</span></span>

<span data-ttu-id="2141d-214">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="2141d-215">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-216">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2141d-217">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="2141d-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-218">Type</span></span>

*   [<span data-ttu-id="2141d-219">Destinatários</span><span class="sxs-lookup"><span data-stu-id="2141d-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-220">Requirements</span></span>

|<span data-ttu-id="2141d-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-221">Requirement</span></span>| <span data-ttu-id="2141d-222">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-224">1.1</span><span class="sxs-lookup"><span data-stu-id="2141d-224">1.1</span></span>|
|[<span data-ttu-id="2141d-225">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-226">ReadItem</span></span>|
|[<span data-ttu-id="2141d-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-228">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="2141d-230">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-231">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="2141d-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-232">Type</span></span>

*   [<span data-ttu-id="2141d-233">Body</span><span class="sxs-lookup"><span data-stu-id="2141d-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-234">Requirements</span></span>

|<span data-ttu-id="2141d-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-235">Requirement</span></span>| <span data-ttu-id="2141d-236">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-238">1.1</span><span class="sxs-lookup"><span data-stu-id="2141d-238">1.1</span></span>|
|[<span data-ttu-id="2141d-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-240">ReadItem</span></span>|
|[<span data-ttu-id="2141d-241">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-243">Example</span></span>

<span data-ttu-id="2141d-244">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="2141d-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2141d-245">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="2141d-246">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="2141d-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-247">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2141d-248">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-249">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-249">Read mode</span></span>

<span data-ttu-id="2141d-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="2141d-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-252">Compose mode</span></span>

<span data-ttu-id="2141d-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

<br>

---
---

##### <a name="type"></a><span data-ttu-id="2141d-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-254">Type</span></span>

*   <span data-ttu-id="2141d-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-256">Requirements</span></span>

|<span data-ttu-id="2141d-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-257">Requirement</span></span>| <span data-ttu-id="2141d-258">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-260">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-260">1.0</span></span>|
|[<span data-ttu-id="2141d-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-262">ReadItem</span></span>|
|[<span data-ttu-id="2141d-263">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="2141d-265">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="2141d-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="2141d-266">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="2141d-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2141d-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="2141d-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2141d-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="2141d-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-271">Type</span></span>

*   <span data-ttu-id="2141d-272">String</span><span class="sxs-lookup"><span data-stu-id="2141d-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-273">Requirements</span></span>

|<span data-ttu-id="2141d-274">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-274">Requirement</span></span>| <span data-ttu-id="2141d-275">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-277">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-277">1.0</span></span>|
|[<span data-ttu-id="2141d-278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-279">ReadItem</span></span>|
|[<span data-ttu-id="2141d-280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-281">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="2141d-283">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="2141d-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="2141d-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-286">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-286">Type</span></span>

*   <span data-ttu-id="2141d-287">Data</span><span class="sxs-lookup"><span data-stu-id="2141d-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-288">Requirements</span></span>

|<span data-ttu-id="2141d-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-289">Requirement</span></span>| <span data-ttu-id="2141d-290">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-292">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-292">1.0</span></span>|
|[<span data-ttu-id="2141d-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-294">ReadItem</span></span>|
|[<span data-ttu-id="2141d-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-296">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="2141d-298">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="2141d-298">dateTimeModified: Date</span></span>

<span data-ttu-id="2141d-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-301">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-302">Type</span></span>

*   <span data-ttu-id="2141d-303">Data</span><span class="sxs-lookup"><span data-stu-id="2141d-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-304">Requirements</span></span>

|<span data-ttu-id="2141d-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-305">Requirement</span></span>| <span data-ttu-id="2141d-306">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-308">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-308">1.0</span></span>|
|[<span data-ttu-id="2141d-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-310">ReadItem</span></span>|
|[<span data-ttu-id="2141d-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-312">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="2141d-314">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="2141d-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2141d-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="2141d-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-318">Read mode</span></span>

<span data-ttu-id="2141d-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="2141d-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-320">Compose mode</span></span>

<span data-ttu-id="2141d-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2141d-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2141d-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="2141d-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2141d-323">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2141d-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2141d-324">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-324">Type</span></span>

*   <span data-ttu-id="2141d-325">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-326">Requirements</span></span>

|<span data-ttu-id="2141d-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-327">Requirement</span></span>| <span data-ttu-id="2141d-328">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-330">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-330">1.0</span></span>|
|[<span data-ttu-id="2141d-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-332">ReadItem</span></span>|
|[<span data-ttu-id="2141d-333">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="2141d-335">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="2141d-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="2141d-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-340">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2141d-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-341">Type</span></span>

*   [<span data-ttu-id="2141d-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2141d-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-343">Requirements</span></span>

|<span data-ttu-id="2141d-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-344">Requirement</span></span>| <span data-ttu-id="2141d-345">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-347">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-347">1.0</span></span>|
|[<span data-ttu-id="2141d-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-349">ReadItem</span></span>|
|[<span data-ttu-id="2141d-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-351">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="2141d-353">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2141d-353">internetMessageId: String</span></span>

<span data-ttu-id="2141d-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-356">Type</span></span>

*   <span data-ttu-id="2141d-357">String</span><span class="sxs-lookup"><span data-stu-id="2141d-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-358">Requirements</span></span>

|<span data-ttu-id="2141d-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-359">Requirement</span></span>| <span data-ttu-id="2141d-360">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-362">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-362">1.0</span></span>|
|[<span data-ttu-id="2141d-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-364">ReadItem</span></span>|
|[<span data-ttu-id="2141d-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-366">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="2141d-368">doclass: String</span><span class="sxs-lookup"><span data-stu-id="2141d-368">itemClass: String</span></span>

<span data-ttu-id="2141d-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2141d-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="2141d-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-373">Type</span></span> | <span data-ttu-id="2141d-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-374">Description</span></span> | <span data-ttu-id="2141d-375">classe de item</span><span class="sxs-lookup"><span data-stu-id="2141d-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="2141d-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="2141d-376">Appointment items</span></span> | <span data-ttu-id="2141d-377">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="2141d-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="2141d-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="2141d-378">Message items</span></span> | <span data-ttu-id="2141d-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="2141d-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="2141d-380">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="2141d-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-381">Type</span></span>

*   <span data-ttu-id="2141d-382">String</span><span class="sxs-lookup"><span data-stu-id="2141d-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-383">Requirements</span></span>

|<span data-ttu-id="2141d-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-384">Requirement</span></span>| <span data-ttu-id="2141d-385">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-387">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-387">1.0</span></span>|
|[<span data-ttu-id="2141d-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-389">ReadItem</span></span>|
|[<span data-ttu-id="2141d-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-391">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2141d-393">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="2141d-393">(nullable) itemId: String</span></span>

<span data-ttu-id="2141d-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="2141d-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2141d-397">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2141d-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2141d-398">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="2141d-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2141d-399">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="2141d-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2141d-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-402">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-402">Type</span></span>

*   <span data-ttu-id="2141d-403">String</span><span class="sxs-lookup"><span data-stu-id="2141d-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-404">Requirements</span></span>

|<span data-ttu-id="2141d-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-405">Requirement</span></span>| <span data-ttu-id="2141d-406">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-407">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-408">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-408">1.0</span></span>|
|[<span data-ttu-id="2141d-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-410">ReadItem</span></span>|
|[<span data-ttu-id="2141d-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-412">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-413">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-413">Example</span></span>

<span data-ttu-id="2141d-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="2141d-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="2141d-416">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-417">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="2141d-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2141d-418">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-419">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-419">Type</span></span>

*   [<span data-ttu-id="2141d-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2141d-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-421">Requirements</span></span>

|<span data-ttu-id="2141d-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-422">Requirement</span></span>| <span data-ttu-id="2141d-423">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-425">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-425">1.0</span></span>|
|[<span data-ttu-id="2141d-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-427">ReadItem</span></span>|
|[<span data-ttu-id="2141d-428">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-429">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="2141d-431">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-432">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-433">Read mode</span></span>

<span data-ttu-id="2141d-434">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-435">Compose mode</span></span>

<span data-ttu-id="2141d-436">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2141d-437">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-437">Type</span></span>

*   <span data-ttu-id="2141d-438">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-439">Requirements</span></span>

|<span data-ttu-id="2141d-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-440">Requirement</span></span>| <span data-ttu-id="2141d-441">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-443">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-443">1.0</span></span>|
|[<span data-ttu-id="2141d-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-445">ReadItem</span></span>|
|[<span data-ttu-id="2141d-446">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-447">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2141d-448">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2141d-448">normalizedSubject: String</span></span>

<span data-ttu-id="2141d-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2141d-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="2141d-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-453">Type</span></span>

*   <span data-ttu-id="2141d-454">String</span><span class="sxs-lookup"><span data-stu-id="2141d-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-455">Requirements</span></span>

|<span data-ttu-id="2141d-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-456">Requirement</span></span>| <span data-ttu-id="2141d-457">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-459">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-459">1.0</span></span>|
|[<span data-ttu-id="2141d-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-461">ReadItem</span></span>|
|[<span data-ttu-id="2141d-462">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-463">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-464">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="2141d-465">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-466">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="2141d-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-467">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-467">Type</span></span>

*   [<span data-ttu-id="2141d-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2141d-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-469">Requirements</span></span>

|<span data-ttu-id="2141d-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-470">Requirement</span></span>| <span data-ttu-id="2141d-471">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-473">1.3</span><span class="sxs-lookup"><span data-stu-id="2141d-473">1.3</span></span>|
|[<span data-ttu-id="2141d-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-475">ReadItem</span></span>|
|[<span data-ttu-id="2141d-476">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-477">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="2141d-479">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="2141d-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-480">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="2141d-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2141d-481">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-482">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-482">Read mode</span></span>

<span data-ttu-id="2141d-483">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="2141d-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-484">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-484">Compose mode</span></span>

<span data-ttu-id="2141d-485">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="2141d-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2141d-486">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-486">Type</span></span>

*   <span data-ttu-id="2141d-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-488">Requirements</span></span>

|<span data-ttu-id="2141d-489">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-489">Requirement</span></span>| <span data-ttu-id="2141d-490">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-491">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-492">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-492">1.0</span></span>|
|[<span data-ttu-id="2141d-493">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-494">ReadItem</span></span>|
|[<span data-ttu-id="2141d-495">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-496">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="2141d-497">organizador: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-500">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-500">Type</span></span>

*   [<span data-ttu-id="2141d-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2141d-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-502">Requirements</span></span>

|<span data-ttu-id="2141d-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-503">Requirement</span></span>| <span data-ttu-id="2141d-504">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-506">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-506">1.0</span></span>|
|[<span data-ttu-id="2141d-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-508">ReadItem</span></span>|
|[<span data-ttu-id="2141d-509">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-510">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-511">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="2141d-512">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="2141d-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-513">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="2141d-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2141d-514">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-515">Read mode</span></span>

<span data-ttu-id="2141d-516">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="2141d-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-517">Compose mode</span></span>

<span data-ttu-id="2141d-518">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="2141d-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2141d-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-519">Type</span></span>

*   <span data-ttu-id="2141d-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-521">Requirements</span></span>

|<span data-ttu-id="2141d-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-522">Requirement</span></span>| <span data-ttu-id="2141d-523">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-525">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-525">1.0</span></span>|
|[<span data-ttu-id="2141d-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-527">ReadItem</span></span>|
|[<span data-ttu-id="2141d-528">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-529">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="2141d-530">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2141d-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2141d-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="2141d-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-535">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2141d-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2141d-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-536">Type</span></span>

*   [<span data-ttu-id="2141d-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2141d-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="2141d-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-538">Requirements</span></span>

|<span data-ttu-id="2141d-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-539">Requirement</span></span>| <span data-ttu-id="2141d-540">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-542">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-542">1.0</span></span>|
|[<span data-ttu-id="2141d-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-544">ReadItem</span></span>|
|[<span data-ttu-id="2141d-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-546">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="2141d-548">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-549">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="2141d-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2141d-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="2141d-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-552">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-552">Read mode</span></span>

<span data-ttu-id="2141d-553">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="2141d-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-554">Compose mode</span></span>

<span data-ttu-id="2141d-555">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2141d-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2141d-556">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="2141d-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2141d-557">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="2141d-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2141d-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-558">Type</span></span>

*   <span data-ttu-id="2141d-559">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-560">Requirements</span></span>

|<span data-ttu-id="2141d-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-561">Requirement</span></span>| <span data-ttu-id="2141d-562">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-564">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-564">1.0</span></span>|
|[<span data-ttu-id="2141d-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-566">ReadItem</span></span>|
|[<span data-ttu-id="2141d-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="2141d-569">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="2141d-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2141d-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="2141d-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-572">Read mode</span></span>

<span data-ttu-id="2141d-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="2141d-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-575">Compose mode</span></span>

<span data-ttu-id="2141d-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="2141d-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2141d-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-577">Type</span></span>

*   <span data-ttu-id="2141d-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-579">Requirements</span></span>

|<span data-ttu-id="2141d-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-580">Requirement</span></span>| <span data-ttu-id="2141d-581">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-583">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-583">1.0</span></span>|
|[<span data-ttu-id="2141d-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-585">ReadItem</span></span>|
|[<span data-ttu-id="2141d-586">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-587">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="2141d-588">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.3) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2141d-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="2141d-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2141d-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2141d-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="2141d-591">Read mode</span></span>

<span data-ttu-id="2141d-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="2141d-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2141d-594">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="2141d-594">Compose mode</span></span>

<span data-ttu-id="2141d-595">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2141d-596">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-596">Type</span></span>

*   <span data-ttu-id="2141d-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-598">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-598">Requirements</span></span>

|<span data-ttu-id="2141d-599">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-599">Requirement</span></span>| <span data-ttu-id="2141d-600">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-601">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-602">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-602">1.0</span></span>|
|[<span data-ttu-id="2141d-603">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-604">ReadItem</span></span>|
|[<span data-ttu-id="2141d-605">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-606">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2141d-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="2141d-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2141d-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2141d-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2141d-609">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="2141d-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2141d-610">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="2141d-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2141d-611">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="2141d-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-612">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-612">Parameters</span></span>

|<span data-ttu-id="2141d-613">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-613">Name</span></span>| <span data-ttu-id="2141d-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-614">Type</span></span>| <span data-ttu-id="2141d-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-615">Attributes</span></span>| <span data-ttu-id="2141d-616">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="2141d-617">String</span><span class="sxs-lookup"><span data-stu-id="2141d-617">String</span></span>||<span data-ttu-id="2141d-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2141d-620">String</span><span class="sxs-lookup"><span data-stu-id="2141d-620">String</span></span>||<span data-ttu-id="2141d-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2141d-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-623">Object</span></span>| <span data-ttu-id="2141d-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-624">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-626">Object</span></span>| <span data-ttu-id="2141d-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-627">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2141d-629">function</span><span class="sxs-lookup"><span data-stu-id="2141d-629">function</span></span>| <span data-ttu-id="2141d-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-630">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-631">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2141d-632">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2141d-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2141d-633">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="2141d-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2141d-634">Erros</span><span class="sxs-lookup"><span data-stu-id="2141d-634">Errors</span></span>

| <span data-ttu-id="2141d-635">Código de erro</span><span class="sxs-lookup"><span data-stu-id="2141d-635">Error code</span></span> | <span data-ttu-id="2141d-636">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="2141d-637">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="2141d-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="2141d-638">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="2141d-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2141d-639">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="2141d-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-640">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-640">Requirements</span></span>

|<span data-ttu-id="2141d-641">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-641">Requirement</span></span>| <span data-ttu-id="2141d-642">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-643">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-644">1.1</span><span class="sxs-lookup"><span data-stu-id="2141d-644">1.1</span></span>|
|[<span data-ttu-id="2141d-645">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-647">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-648">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-649">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-649">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2141d-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2141d-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2141d-651">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2141d-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="2141d-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2141d-655">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="2141d-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2141d-656">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="2141d-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-657">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-657">Parameters</span></span>

|<span data-ttu-id="2141d-658">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-658">Name</span></span>| <span data-ttu-id="2141d-659">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-659">Type</span></span>| <span data-ttu-id="2141d-660">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-660">Attributes</span></span>| <span data-ttu-id="2141d-661">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="2141d-662">String</span><span class="sxs-lookup"><span data-stu-id="2141d-662">String</span></span>||<span data-ttu-id="2141d-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2141d-665">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2141d-665">String</span></span>||<span data-ttu-id="2141d-666">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="2141d-666">The subject of the item to be attached.</span></span> <span data-ttu-id="2141d-667">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2141d-668">Object</span><span class="sxs-lookup"><span data-stu-id="2141d-668">Object</span></span>| <span data-ttu-id="2141d-669">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-669">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-670">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-671">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-671">Object</span></span>| <span data-ttu-id="2141d-672">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-672">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-673">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2141d-674">function</span><span class="sxs-lookup"><span data-stu-id="2141d-674">function</span></span>| <span data-ttu-id="2141d-675">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-675">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-676">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2141d-677">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2141d-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2141d-678">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="2141d-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2141d-679">Erros</span><span class="sxs-lookup"><span data-stu-id="2141d-679">Errors</span></span>

| <span data-ttu-id="2141d-680">Código de erro</span><span class="sxs-lookup"><span data-stu-id="2141d-680">Error code</span></span> | <span data-ttu-id="2141d-681">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2141d-682">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="2141d-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-683">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-683">Requirements</span></span>

|<span data-ttu-id="2141d-684">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-684">Requirement</span></span>| <span data-ttu-id="2141d-685">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-686">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-687">1.1</span><span class="sxs-lookup"><span data-stu-id="2141d-687">1.1</span></span>|
|[<span data-ttu-id="2141d-688">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-690">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-691">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-692">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-692">Example</span></span>

<span data-ttu-id="2141d-693">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="2141d-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="2141d-694">close()</span><span class="sxs-lookup"><span data-stu-id="2141d-694">close()</span></span>

<span data-ttu-id="2141d-695">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="2141d-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2141d-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="2141d-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-698">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="2141d-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2141d-699">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="2141d-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-700">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-700">Requirements</span></span>

|<span data-ttu-id="2141d-701">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-701">Requirement</span></span>| <span data-ttu-id="2141d-702">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-703">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-704">1.3</span><span class="sxs-lookup"><span data-stu-id="2141d-704">1.3</span></span>|
|[<span data-ttu-id="2141d-705">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-706">Restrito</span><span class="sxs-lookup"><span data-stu-id="2141d-706">Restricted</span></span>|
|[<span data-ttu-id="2141d-707">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-708">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-708">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2141d-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2141d-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2141d-710">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="2141d-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-711">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2141d-712">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="2141d-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2141d-713">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="2141d-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2141d-714">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="2141d-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="2141d-715">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="2141d-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="2141d-716">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="2141d-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-717">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-717">Parameters</span></span>

|<span data-ttu-id="2141d-718">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-718">Name</span></span>| <span data-ttu-id="2141d-719">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-719">Type</span></span>| <span data-ttu-id="2141d-720">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2141d-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2141d-721">String &#124; Object</span></span>| |<span data-ttu-id="2141d-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2141d-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2141d-724">**OU**</span><span class="sxs-lookup"><span data-stu-id="2141d-724">**OR**</span></span><br/><span data-ttu-id="2141d-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2141d-727">String</span><span class="sxs-lookup"><span data-stu-id="2141d-727">String</span></span> | <span data-ttu-id="2141d-728">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-728">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2141d-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2141d-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2141d-732">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-732">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-733">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="2141d-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2141d-734">String</span><span class="sxs-lookup"><span data-stu-id="2141d-734">String</span></span> | | <span data-ttu-id="2141d-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="2141d-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2141d-737">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2141d-737">String</span></span> | | <span data-ttu-id="2141d-738">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="2141d-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2141d-739">String</span><span class="sxs-lookup"><span data-stu-id="2141d-739">String</span></span> | | <span data-ttu-id="2141d-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="2141d-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2141d-742">String</span><span class="sxs-lookup"><span data-stu-id="2141d-742">String</span></span> | | <span data-ttu-id="2141d-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2141d-746">function</span><span class="sxs-lookup"><span data-stu-id="2141d-746">function</span></span> | <span data-ttu-id="2141d-747">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-747">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-748">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-749">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-749">Requirements</span></span>

|<span data-ttu-id="2141d-750">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-750">Requirement</span></span>| <span data-ttu-id="2141d-751">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-752">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-753">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-753">1.0</span></span>|
|[<span data-ttu-id="2141d-754">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-755">ReadItem</span></span>|
|[<span data-ttu-id="2141d-756">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-757">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2141d-758">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2141d-758">Examples</span></span>

<span data-ttu-id="2141d-759">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="2141d-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2141d-760">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="2141d-760">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2141d-761">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="2141d-761">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2141d-762">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="2141d-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2141d-763">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="2141d-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2141d-764">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2141d-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2141d-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2141d-766">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="2141d-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-767">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2141d-768">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="2141d-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2141d-769">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="2141d-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2141d-770">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="2141d-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="2141d-771">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="2141d-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="2141d-772">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="2141d-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-773">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-773">Parameters</span></span>

|<span data-ttu-id="2141d-774">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-774">Name</span></span>| <span data-ttu-id="2141d-775">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-775">Type</span></span>| <span data-ttu-id="2141d-776">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2141d-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2141d-777">String &#124; Object</span></span>| | <span data-ttu-id="2141d-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2141d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2141d-780">**OU**</span><span class="sxs-lookup"><span data-stu-id="2141d-780">**OR**</span></span><br/><span data-ttu-id="2141d-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2141d-783">String</span><span class="sxs-lookup"><span data-stu-id="2141d-783">String</span></span> | <span data-ttu-id="2141d-784">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-784">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2141d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2141d-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2141d-788">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-788">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-789">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="2141d-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2141d-790">String</span><span class="sxs-lookup"><span data-stu-id="2141d-790">String</span></span> | | <span data-ttu-id="2141d-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="2141d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2141d-793">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2141d-793">String</span></span> | | <span data-ttu-id="2141d-794">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="2141d-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2141d-795">String</span><span class="sxs-lookup"><span data-stu-id="2141d-795">String</span></span> | | <span data-ttu-id="2141d-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="2141d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2141d-798">String</span><span class="sxs-lookup"><span data-stu-id="2141d-798">String</span></span> | | <span data-ttu-id="2141d-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2141d-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2141d-802">function</span><span class="sxs-lookup"><span data-stu-id="2141d-802">function</span></span> | <span data-ttu-id="2141d-803">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-803">&lt;optional&gt;</span></span> | <span data-ttu-id="2141d-804">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-805">Requirements</span></span>

|<span data-ttu-id="2141d-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-806">Requirement</span></span>| <span data-ttu-id="2141d-807">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-808">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-809">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-809">1.0</span></span>|
|[<span data-ttu-id="2141d-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-811">ReadItem</span></span>|
|[<span data-ttu-id="2141d-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-813">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2141d-814">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2141d-814">Examples</span></span>

<span data-ttu-id="2141d-815">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="2141d-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2141d-816">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="2141d-816">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2141d-817">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="2141d-817">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2141d-818">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="2141d-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2141d-819">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="2141d-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2141d-820">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="2141d-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="2141d-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="2141d-822">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="2141d-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-823">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-824">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-824">Requirements</span></span>

|<span data-ttu-id="2141d-825">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-825">Requirement</span></span>| <span data-ttu-id="2141d-826">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-827">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-828">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-828">1.0</span></span>|
|[<span data-ttu-id="2141d-829">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-830">ReadItem</span></span>|
|[<span data-ttu-id="2141d-831">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-832">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-833">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-833">Returns:</span></span>

<span data-ttu-id="2141d-834">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="2141d-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="2141d-835">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-835">Example</span></span>

<span data-ttu-id="2141d-836">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-836">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="2141d-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="2141d-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="2141d-838">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="2141d-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-839">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-840">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-840">Parameters</span></span>

|<span data-ttu-id="2141d-841">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-841">Name</span></span>| <span data-ttu-id="2141d-842">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-842">Type</span></span>| <span data-ttu-id="2141d-843">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="2141d-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2141d-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="2141d-845">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="2141d-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-846">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-846">Requirements</span></span>

|<span data-ttu-id="2141d-847">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-847">Requirement</span></span>| <span data-ttu-id="2141d-848">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-849">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-850">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-850">1.0</span></span>|
|[<span data-ttu-id="2141d-851">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-852">Restrito</span><span class="sxs-lookup"><span data-stu-id="2141d-852">Restricted</span></span>|
|[<span data-ttu-id="2141d-853">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-854">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-855">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-855">Returns:</span></span>

<span data-ttu-id="2141d-856">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="2141d-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2141d-857">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="2141d-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2141d-858">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="2141d-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2141d-859">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="2141d-860">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="2141d-860">Value of `entityType`</span></span> | <span data-ttu-id="2141d-861">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="2141d-861">Type of objects in returned array</span></span> | <span data-ttu-id="2141d-862">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="2141d-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="2141d-863">String</span><span class="sxs-lookup"><span data-stu-id="2141d-863">String</span></span> | <span data-ttu-id="2141d-864">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="2141d-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="2141d-865">Contato</span><span class="sxs-lookup"><span data-stu-id="2141d-865">Contact</span></span> | <span data-ttu-id="2141d-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2141d-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="2141d-867">String</span><span class="sxs-lookup"><span data-stu-id="2141d-867">String</span></span> | <span data-ttu-id="2141d-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2141d-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="2141d-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2141d-869">MeetingSuggestion</span></span> | <span data-ttu-id="2141d-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2141d-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="2141d-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2141d-871">PhoneNumber</span></span> | <span data-ttu-id="2141d-872">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="2141d-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="2141d-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2141d-873">TaskSuggestion</span></span> | <span data-ttu-id="2141d-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2141d-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="2141d-875">String</span><span class="sxs-lookup"><span data-stu-id="2141d-875">String</span></span> | <span data-ttu-id="2141d-876">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="2141d-876">**Restricted**</span></span> |

<span data-ttu-id="2141d-877">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="2141d-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="2141d-878">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-878">Example</span></span>

<span data-ttu-id="2141d-879">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="2141d-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="2141d-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="2141d-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="2141d-881">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="2141d-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-882">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2141d-883">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="2141d-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-884">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-884">Parameters</span></span>

|<span data-ttu-id="2141d-885">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-885">Name</span></span>| <span data-ttu-id="2141d-886">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-886">Type</span></span>| <span data-ttu-id="2141d-887">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2141d-888">String</span><span class="sxs-lookup"><span data-stu-id="2141d-888">String</span></span>|<span data-ttu-id="2141d-889">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="2141d-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-890">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-890">Requirements</span></span>

|<span data-ttu-id="2141d-891">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-891">Requirement</span></span>| <span data-ttu-id="2141d-892">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-893">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-894">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-894">1.0</span></span>|
|[<span data-ttu-id="2141d-895">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-896">ReadItem</span></span>|
|[<span data-ttu-id="2141d-897">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-898">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-899">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-899">Returns:</span></span>

<span data-ttu-id="2141d-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="2141d-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2141d-902">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="2141d-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="2141d-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2141d-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2141d-904">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="2141d-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-905">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2141d-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="2141d-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2141d-909">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="2141d-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2141d-910">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2141d-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2141d-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="2141d-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2141d-914">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-914">Requirements</span></span>

|<span data-ttu-id="2141d-915">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-915">Requirement</span></span>| <span data-ttu-id="2141d-916">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-917">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-918">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-918">1.0</span></span>|
|[<span data-ttu-id="2141d-919">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-920">ReadItem</span></span>|
|[<span data-ttu-id="2141d-921">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-922">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-923">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-923">Returns:</span></span>

<span data-ttu-id="2141d-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="2141d-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="2141d-926">Tipo: objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-926">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="2141d-927">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-927">Example</span></span>

<span data-ttu-id="2141d-928">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="2141d-928">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2141d-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="2141d-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2141d-930">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="2141d-930">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-931">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2141d-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2141d-932">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="2141d-932">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2141d-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="2141d-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-935">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-935">Parameters</span></span>

|<span data-ttu-id="2141d-936">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-936">Name</span></span>| <span data-ttu-id="2141d-937">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-937">Type</span></span>| <span data-ttu-id="2141d-938">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-938">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2141d-939">String</span><span class="sxs-lookup"><span data-stu-id="2141d-939">String</span></span>|<span data-ttu-id="2141d-940">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="2141d-940">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-941">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-941">Requirements</span></span>

|<span data-ttu-id="2141d-942">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-942">Requirement</span></span>| <span data-ttu-id="2141d-943">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-944">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-945">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-945">1.0</span></span>|
|[<span data-ttu-id="2141d-946">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-947">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-947">ReadItem</span></span>|
|[<span data-ttu-id="2141d-948">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-949">Read</span><span class="sxs-lookup"><span data-stu-id="2141d-949">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-950">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-950">Returns:</span></span>

<span data-ttu-id="2141d-951">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="2141d-951">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="2141d-952">Tipo: cadeia de caracteres de matriz. < ></span><span class="sxs-lookup"><span data-stu-id="2141d-952">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="2141d-953">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-953">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2141d-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2141d-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2141d-955">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-955">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2141d-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="2141d-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-958">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-958">Parameters</span></span>

|<span data-ttu-id="2141d-959">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-959">Name</span></span>| <span data-ttu-id="2141d-960">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-960">Type</span></span>| <span data-ttu-id="2141d-961">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-961">Attributes</span></span>| <span data-ttu-id="2141d-962">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-962">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="2141d-963">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2141d-963">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2141d-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="2141d-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="2141d-967">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-967">Object</span></span>| <span data-ttu-id="2141d-968">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-968">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-969">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-969">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-970">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-970">Object</span></span>| <span data-ttu-id="2141d-971">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-971">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-972">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-972">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2141d-973">function</span><span class="sxs-lookup"><span data-stu-id="2141d-973">function</span></span>||<span data-ttu-id="2141d-974">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-974">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2141d-975">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="2141d-975">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2141d-976">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="2141d-976">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-977">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-977">Requirements</span></span>

|<span data-ttu-id="2141d-978">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-978">Requirement</span></span>| <span data-ttu-id="2141d-979">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-979">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-980">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-980">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-981">1.2</span><span class="sxs-lookup"><span data-stu-id="2141d-981">1.2</span></span>|
|[<span data-ttu-id="2141d-982">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-982">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-983">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-983">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-984">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-984">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-985">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-985">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2141d-986">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2141d-986">Returns:</span></span>

<span data-ttu-id="2141d-987">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="2141d-987">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="2141d-988">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="2141d-988">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2141d-989">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-989">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2141d-990">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2141d-990">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2141d-991">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="2141d-991">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2141d-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="2141d-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-995">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-995">Parameters</span></span>

|<span data-ttu-id="2141d-996">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-996">Name</span></span>| <span data-ttu-id="2141d-997">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-997">Type</span></span>| <span data-ttu-id="2141d-998">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-998">Attributes</span></span>| <span data-ttu-id="2141d-999">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-999">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2141d-1000">function</span><span class="sxs-lookup"><span data-stu-id="2141d-1000">function</span></span>||<span data-ttu-id="2141d-1001">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-1001">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2141d-1002">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2141d-1002">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2141d-1003">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="2141d-1003">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="2141d-1004">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1004">Object</span></span>| <span data-ttu-id="2141d-1005">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1005">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1006">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1006">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2141d-1007">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1007">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-1008">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-1008">Requirements</span></span>

|<span data-ttu-id="2141d-1009">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-1009">Requirement</span></span>| <span data-ttu-id="2141d-1010">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-1010">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-1011">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-1011">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-1012">1.0</span><span class="sxs-lookup"><span data-stu-id="2141d-1012">1.0</span></span>|
|[<span data-ttu-id="2141d-1013">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-1013">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-1014">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2141d-1014">ReadItem</span></span>|
|[<span data-ttu-id="2141d-1015">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2141d-1015">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-1016">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2141d-1016">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-1017">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-1017">Example</span></span>

<span data-ttu-id="2141d-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="2141d-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2141d-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2141d-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2141d-1022">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="2141d-1022">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2141d-1023">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="2141d-1023">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="2141d-1024">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="2141d-1024">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="2141d-1025">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="2141d-1025">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="2141d-1026">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1026">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-1027">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-1027">Parameters</span></span>

|<span data-ttu-id="2141d-1028">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-1028">Name</span></span>| <span data-ttu-id="2141d-1029">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-1029">Type</span></span>| <span data-ttu-id="2141d-1030">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-1030">Attributes</span></span>| <span data-ttu-id="2141d-1031">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-1031">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="2141d-1032">String</span><span class="sxs-lookup"><span data-stu-id="2141d-1032">String</span></span>||<span data-ttu-id="2141d-1033">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="2141d-1033">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="2141d-1034">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1034">Object</span></span>| <span data-ttu-id="2141d-1035">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1036">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-1036">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-1037">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1037">Object</span></span>| <span data-ttu-id="2141d-1038">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1039">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1039">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2141d-1040">function</span><span class="sxs-lookup"><span data-stu-id="2141d-1040">function</span></span>| <span data-ttu-id="2141d-1041">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1042">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-1042">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2141d-1043">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="2141d-1043">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2141d-1044">Erros</span><span class="sxs-lookup"><span data-stu-id="2141d-1044">Errors</span></span>

| <span data-ttu-id="2141d-1045">Código de erro</span><span class="sxs-lookup"><span data-stu-id="2141d-1045">Error code</span></span> | <span data-ttu-id="2141d-1046">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-1046">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="2141d-1047">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="2141d-1047">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-1048">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-1048">Requirements</span></span>

|<span data-ttu-id="2141d-1049">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-1049">Requirement</span></span>| <span data-ttu-id="2141d-1050">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-1051">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-1052">1.1</span><span class="sxs-lookup"><span data-stu-id="2141d-1052">1.1</span></span>|
|[<span data-ttu-id="2141d-1053">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-1054">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-1054">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-1055">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-1056">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-1056">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-1057">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-1057">Example</span></span>

<span data-ttu-id="2141d-1058">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="2141d-1058">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="2141d-1059">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2141d-1059">saveAsync([options], callback)</span></span>

<span data-ttu-id="2141d-1060">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="2141d-1060">Asynchronously saves an item.</span></span>

<span data-ttu-id="2141d-1061">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1061">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="2141d-1062">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="2141d-1062">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="2141d-1063">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="2141d-1063">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-1064">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="2141d-1064">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2141d-1065">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="2141d-1065">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2141d-p168">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="2141d-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2141d-1069">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="2141d-1069">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2141d-1070">O Outlook no Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="2141d-1070">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="2141d-1071">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="2141d-1071">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="2141d-1072">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="2141d-1072">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="2141d-1073">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="2141d-1073">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-1074">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-1074">Parameters</span></span>

|<span data-ttu-id="2141d-1075">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-1075">Name</span></span>| <span data-ttu-id="2141d-1076">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-1076">Type</span></span>| <span data-ttu-id="2141d-1077">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-1077">Attributes</span></span>| <span data-ttu-id="2141d-1078">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-1078">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="2141d-1079">Object</span><span class="sxs-lookup"><span data-stu-id="2141d-1079">Object</span></span>| <span data-ttu-id="2141d-1080">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1080">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1081">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-1081">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-1082">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1082">Object</span></span>| <span data-ttu-id="2141d-1083">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1084">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1084">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2141d-1085">function</span><span class="sxs-lookup"><span data-stu-id="2141d-1085">function</span></span>||<span data-ttu-id="2141d-1086">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2141d-1087">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2141d-1087">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2141d-1088">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-1088">Requirements</span></span>

|<span data-ttu-id="2141d-1089">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-1089">Requirement</span></span>| <span data-ttu-id="2141d-1090">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-1091">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-1092">1.3</span><span class="sxs-lookup"><span data-stu-id="2141d-1092">1.3</span></span>|
|[<span data-ttu-id="2141d-1093">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-1095">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-1096">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-1096">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2141d-1097">Exemplos</span><span class="sxs-lookup"><span data-stu-id="2141d-1097">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2141d-p170">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="2141d-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2141d-1100">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2141d-1100">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2141d-1101">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2141d-1101">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2141d-p171">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="2141d-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2141d-1105">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2141d-1105">Parameters</span></span>

|<span data-ttu-id="2141d-1106">Nome</span><span class="sxs-lookup"><span data-stu-id="2141d-1106">Name</span></span>| <span data-ttu-id="2141d-1107">Tipo</span><span class="sxs-lookup"><span data-stu-id="2141d-1107">Type</span></span>| <span data-ttu-id="2141d-1108">Atributos</span><span class="sxs-lookup"><span data-stu-id="2141d-1108">Attributes</span></span>| <span data-ttu-id="2141d-1109">Descrição</span><span class="sxs-lookup"><span data-stu-id="2141d-1109">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2141d-1110">String</span><span class="sxs-lookup"><span data-stu-id="2141d-1110">String</span></span>||<span data-ttu-id="2141d-p172">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="2141d-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="2141d-1114">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1114">Object</span></span>| <span data-ttu-id="2141d-1115">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1115">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1116">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="2141d-1116">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2141d-1117">Objeto</span><span class="sxs-lookup"><span data-stu-id="2141d-1117">Object</span></span>| <span data-ttu-id="2141d-1118">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1119">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="2141d-1119">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2141d-1120">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2141d-1120">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2141d-1121">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2141d-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="2141d-1122">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2141d-1122">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="2141d-1123">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="2141d-1123">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2141d-1124">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2141d-1124">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="2141d-1125">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="2141d-1125">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2141d-1126">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="2141d-1126">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="2141d-1127">function</span><span class="sxs-lookup"><span data-stu-id="2141d-1127">function</span></span>||<span data-ttu-id="2141d-1128">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2141d-1128">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2141d-1129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2141d-1129">Requirements</span></span>

|<span data-ttu-id="2141d-1130">Requisito</span><span class="sxs-lookup"><span data-stu-id="2141d-1130">Requirement</span></span>| <span data-ttu-id="2141d-1131">Valor</span><span class="sxs-lookup"><span data-stu-id="2141d-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="2141d-1132">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2141d-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2141d-1133">1.2</span><span class="sxs-lookup"><span data-stu-id="2141d-1133">1.2</span></span>|
|[<span data-ttu-id="2141d-1134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2141d-1134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2141d-1135">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2141d-1135">ReadWriteItem</span></span>|
|[<span data-ttu-id="2141d-1136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2141d-1136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2141d-1137">Escrever</span><span class="sxs-lookup"><span data-stu-id="2141d-1137">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2141d-1138">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2141d-1138">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
