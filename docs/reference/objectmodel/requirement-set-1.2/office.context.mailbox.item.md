---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,2
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 97fa271f500e89c6ce69d82b95a0818f6d5bc7d4
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001604"
---
# <a name="item"></a><span data-ttu-id="8ab66-102">item</span><span class="sxs-lookup"><span data-stu-id="8ab66-102">item</span></span>

### <span data-ttu-id="8ab66-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="8ab66-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="8ab66-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8ab66-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-107">Requirements</span></span>

|<span data-ttu-id="8ab66-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-108">Requirement</span></span>| <span data-ttu-id="8ab66-109">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-111">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-111">1.0</span></span>|
|[<span data-ttu-id="8ab66-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="8ab66-113">Restricted</span></span>|
|[<span data-ttu-id="8ab66-114">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-115">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8ab66-116">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8ab66-116">Members and methods</span></span>

| <span data-ttu-id="8ab66-117">Membro	</span><span class="sxs-lookup"><span data-stu-id="8ab66-117">Member</span></span> | <span data-ttu-id="8ab66-118">Tipo	</span><span class="sxs-lookup"><span data-stu-id="8ab66-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8ab66-119">attachments</span><span class="sxs-lookup"><span data-stu-id="8ab66-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="8ab66-120">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-120">Member</span></span> |
| [<span data-ttu-id="8ab66-121">bcc</span><span class="sxs-lookup"><span data-stu-id="8ab66-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="8ab66-122">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-122">Member</span></span> |
| [<span data-ttu-id="8ab66-123">body</span><span class="sxs-lookup"><span data-stu-id="8ab66-123">body</span></span>](#body-body) | <span data-ttu-id="8ab66-124">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-124">Member</span></span> |
| [<span data-ttu-id="8ab66-125">cc</span><span class="sxs-lookup"><span data-stu-id="8ab66-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ab66-126">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-126">Member</span></span> |
| [<span data-ttu-id="8ab66-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="8ab66-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8ab66-128">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-128">Member</span></span> |
| [<span data-ttu-id="8ab66-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8ab66-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8ab66-130">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-130">Member</span></span> |
| [<span data-ttu-id="8ab66-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8ab66-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8ab66-132">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-132">Member</span></span> |
| [<span data-ttu-id="8ab66-133">end</span><span class="sxs-lookup"><span data-stu-id="8ab66-133">end</span></span>](#end-datetime) | <span data-ttu-id="8ab66-134">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-134">Member</span></span> |
| [<span data-ttu-id="8ab66-135">from</span><span class="sxs-lookup"><span data-stu-id="8ab66-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="8ab66-136">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-136">Member</span></span> |
| [<span data-ttu-id="8ab66-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8ab66-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8ab66-138">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-138">Member</span></span> |
| [<span data-ttu-id="8ab66-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="8ab66-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8ab66-140">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-140">Member</span></span> |
| [<span data-ttu-id="8ab66-141">itemId</span><span class="sxs-lookup"><span data-stu-id="8ab66-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8ab66-142">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-142">Member</span></span> |
| [<span data-ttu-id="8ab66-143">itemType</span><span class="sxs-lookup"><span data-stu-id="8ab66-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="8ab66-144">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-144">Member</span></span> |
| [<span data-ttu-id="8ab66-145">location</span><span class="sxs-lookup"><span data-stu-id="8ab66-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="8ab66-146">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-146">Member</span></span> |
| [<span data-ttu-id="8ab66-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8ab66-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8ab66-148">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-148">Member</span></span> |
| [<span data-ttu-id="8ab66-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8ab66-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ab66-150">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-150">Member</span></span> |
| [<span data-ttu-id="8ab66-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8ab66-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="8ab66-152">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-152">Member</span></span> |
| [<span data-ttu-id="8ab66-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8ab66-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ab66-154">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-154">Member</span></span> |
| [<span data-ttu-id="8ab66-155">sender</span><span class="sxs-lookup"><span data-stu-id="8ab66-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="8ab66-156">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-156">Member</span></span> |
| [<span data-ttu-id="8ab66-157">start</span><span class="sxs-lookup"><span data-stu-id="8ab66-157">start</span></span>](#start-datetime) | <span data-ttu-id="8ab66-158">Member</span><span class="sxs-lookup"><span data-stu-id="8ab66-158">Member</span></span> |
| [<span data-ttu-id="8ab66-159">subject</span><span class="sxs-lookup"><span data-stu-id="8ab66-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="8ab66-160">Membro</span><span class="sxs-lookup"><span data-stu-id="8ab66-160">Member</span></span> |
| [<span data-ttu-id="8ab66-161">to</span><span class="sxs-lookup"><span data-stu-id="8ab66-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ab66-162">Membro</span><span class="sxs-lookup"><span data-stu-id="8ab66-162">Member</span></span> |
| [<span data-ttu-id="8ab66-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8ab66-164">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-164">Method</span></span> |
| [<span data-ttu-id="8ab66-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8ab66-166">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-166">Method</span></span> |
| [<span data-ttu-id="8ab66-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8ab66-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="8ab66-168">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-168">Method</span></span> |
| [<span data-ttu-id="8ab66-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8ab66-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="8ab66-170">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-170">Method</span></span> |
| [<span data-ttu-id="8ab66-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="8ab66-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="8ab66-172">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-172">Method</span></span> |
| [<span data-ttu-id="8ab66-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8ab66-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8ab66-174">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-174">Method</span></span> |
| [<span data-ttu-id="8ab66-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8ab66-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8ab66-176">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-176">Method</span></span> |
| [<span data-ttu-id="8ab66-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8ab66-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8ab66-178">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-178">Method</span></span> |
| [<span data-ttu-id="8ab66-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8ab66-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8ab66-180">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-180">Method</span></span> |
| [<span data-ttu-id="8ab66-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8ab66-182">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-182">Method</span></span> |
| [<span data-ttu-id="8ab66-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8ab66-184">Method</span><span class="sxs-lookup"><span data-stu-id="8ab66-184">Method</span></span> |
| [<span data-ttu-id="8ab66-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8ab66-186">Método</span><span class="sxs-lookup"><span data-stu-id="8ab66-186">Method</span></span> |
| [<span data-ttu-id="8ab66-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8ab66-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8ab66-188">Método</span><span class="sxs-lookup"><span data-stu-id="8ab66-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8ab66-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-189">Example</span></span>

<span data-ttu-id="8ab66-190">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8ab66-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8ab66-191">Members</span><span class="sxs-lookup"><span data-stu-id="8ab66-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="8ab66-192">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="8ab66-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="8ab66-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-195">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8ab66-196">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8ab66-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-197">Type</span></span>

*   <span data-ttu-id="8ab66-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="8ab66-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-199">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-199">Requirements</span></span>

|<span data-ttu-id="8ab66-200">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-200">Requirement</span></span>| <span data-ttu-id="8ab66-201">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-202">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-203">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-203">1.0</span></span>|
|[<span data-ttu-id="8ab66-204">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-205">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-206">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-207">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-208">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-208">Example</span></span>

<span data-ttu-id="8ab66-209">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="8ab66-210">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-211">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8ab66-212">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8ab66-212">Compose mode only.</span></span>

<span data-ttu-id="8ab66-213">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-214">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ab66-215">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ab66-216">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="8ab66-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-217">Type</span></span>

*   [<span data-ttu-id="8ab66-218">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8ab66-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-219">Requirements</span></span>

|<span data-ttu-id="8ab66-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-220">Requirement</span></span>| <span data-ttu-id="8ab66-221">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-223">1.1</span><span class="sxs-lookup"><span data-stu-id="8ab66-223">1.1</span></span>|
|[<span data-ttu-id="8ab66-224">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-225">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-227">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-228">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="8ab66-229">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-230">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-231">Type</span></span>

*   [<span data-ttu-id="8ab66-232">Body</span><span class="sxs-lookup"><span data-stu-id="8ab66-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-233">Requirements</span></span>

|<span data-ttu-id="8ab66-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-234">Requirement</span></span>| <span data-ttu-id="8ab66-235">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-237">1.1</span><span class="sxs-lookup"><span data-stu-id="8ab66-237">1.1</span></span>|
|[<span data-ttu-id="8ab66-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-239">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-240">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-241">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-242">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-242">Example</span></span>

<span data-ttu-id="8ab66-243">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="8ab66-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8ab66-244">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="8ab66-245">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-246">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8ab66-247">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-248">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-248">Read mode</span></span>

<span data-ttu-id="8ab66-249">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="8ab66-250">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-251">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-252">Compose mode</span></span>

<span data-ttu-id="8ab66-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="8ab66-254">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-255">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ab66-256">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ab66-257">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="8ab66-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-258">Type</span></span>

*   <span data-ttu-id="8ab66-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-260">Requirements</span></span>

|<span data-ttu-id="8ab66-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-261">Requirement</span></span>| <span data-ttu-id="8ab66-262">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-264">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-264">1.0</span></span>|
|[<span data-ttu-id="8ab66-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-266">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8ab66-269">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="8ab66-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="8ab66-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8ab66-p110">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8ab66-p111">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-275">Type</span></span>

*   <span data-ttu-id="8ab66-276">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-277">Requirements</span></span>

|<span data-ttu-id="8ab66-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-278">Requirement</span></span>| <span data-ttu-id="8ab66-279">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-281">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-281">1.0</span></span>|
|[<span data-ttu-id="8ab66-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-283">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="8ab66-287">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="8ab66-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="8ab66-p112">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-290">Type</span></span>

*   <span data-ttu-id="8ab66-291">Data</span><span class="sxs-lookup"><span data-stu-id="8ab66-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-292">Requirements</span></span>

|<span data-ttu-id="8ab66-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-293">Requirement</span></span>| <span data-ttu-id="8ab66-294">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-296">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-296">1.0</span></span>|
|[<span data-ttu-id="8ab66-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-298">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-300">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="8ab66-302">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="8ab66-302">dateTimeModified: Date</span></span>

<span data-ttu-id="8ab66-p113">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-305">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-306">Type</span></span>

*   <span data-ttu-id="8ab66-307">Data</span><span class="sxs-lookup"><span data-stu-id="8ab66-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-308">Requirements</span></span>

|<span data-ttu-id="8ab66-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-309">Requirement</span></span>| <span data-ttu-id="8ab66-310">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-312">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-312">1.0</span></span>|
|[<span data-ttu-id="8ab66-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-314">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-316">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="8ab66-318">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="8ab66-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8ab66-p114">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-322">Read mode</span></span>

<span data-ttu-id="8ab66-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-324">Compose mode</span></span>

<span data-ttu-id="8ab66-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8ab66-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8ab66-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8ab66-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8ab66-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-328">Type</span></span>

*   <span data-ttu-id="8ab66-329">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-330">Requirements</span></span>

|<span data-ttu-id="8ab66-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-331">Requirement</span></span>| <span data-ttu-id="8ab66-332">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-334">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-334">1.0</span></span>|
|[<span data-ttu-id="8ab66-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-336">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="8ab66-339">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-p115">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8ab66-p116">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-345">Type</span></span>

*   [<span data-ttu-id="8ab66-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ab66-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-347">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-347">Requirements</span></span>

|<span data-ttu-id="8ab66-348">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-348">Requirement</span></span>| <span data-ttu-id="8ab66-349">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-350">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-351">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-351">1.0</span></span>|
|[<span data-ttu-id="8ab66-352">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-353">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-354">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-355">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-356">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="8ab66-357">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-357">internetMessageId: String</span></span>

<span data-ttu-id="8ab66-p117">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-360">Type</span></span>

*   <span data-ttu-id="8ab66-361">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-362">Requirements</span></span>

|<span data-ttu-id="8ab66-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-363">Requirement</span></span>| <span data-ttu-id="8ab66-364">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-366">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-366">1.0</span></span>|
|[<span data-ttu-id="8ab66-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-368">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-370">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="8ab66-372">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="8ab66-372">itemClass: String</span></span>

<span data-ttu-id="8ab66-p118">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8ab66-p119">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8ab66-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-377">Type</span></span> | <span data-ttu-id="8ab66-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-378">Description</span></span> | <span data-ttu-id="8ab66-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="8ab66-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8ab66-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="8ab66-380">Appointment items</span></span> | <span data-ttu-id="8ab66-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8ab66-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="8ab66-382">Message items</span></span> | <span data-ttu-id="8ab66-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="8ab66-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8ab66-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-385">Type</span></span>

*   <span data-ttu-id="8ab66-386">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-387">Requirements</span></span>

|<span data-ttu-id="8ab66-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-388">Requirement</span></span>| <span data-ttu-id="8ab66-389">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-391">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-391">1.0</span></span>|
|[<span data-ttu-id="8ab66-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-393">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-395">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8ab66-397">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-397">(nullable) itemId: String</span></span>

<span data-ttu-id="8ab66-398">Obtém o [identificador do item dos serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-398">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="8ab66-399">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-399">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-400">O identificador retornado pela `itemId` propriedade é o mesmo que o identificador de [item dos serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="8ab66-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="8ab66-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8ab66-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8ab66-402">Antes de fazer chamadas da API REST usando esse valor, ele deve ser `Office.context.mailbox.convertToRestId`convertido usando o, que está disponível a partir do conjunto de requisitos 1,3.</span><span class="sxs-lookup"><span data-stu-id="8ab66-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="8ab66-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8ab66-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-404">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-404">Type</span></span>

*   <span data-ttu-id="8ab66-405">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-406">Requirements</span></span>

|<span data-ttu-id="8ab66-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-407">Requirement</span></span>| <span data-ttu-id="8ab66-408">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-410">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-410">1.0</span></span>|
|[<span data-ttu-id="8ab66-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-412">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-414">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-415">Example</span></span>

<span data-ttu-id="8ab66-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="8ab66-418">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-419">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="8ab66-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8ab66-420">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-421">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-421">Type</span></span>

*   [<span data-ttu-id="8ab66-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8ab66-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-423">Requirements</span></span>

|<span data-ttu-id="8ab66-424">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-424">Requirement</span></span>| <span data-ttu-id="8ab66-425">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-426">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-427">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-427">1.0</span></span>|
|[<span data-ttu-id="8ab66-428">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-429">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-430">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-431">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-432">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="8ab66-433">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-434">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-435">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-435">Read mode</span></span>

<span data-ttu-id="8ab66-436">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-437">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-437">Compose mode</span></span>

<span data-ttu-id="8ab66-438">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-439">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-439">Type</span></span>

*   <span data-ttu-id="8ab66-440">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-441">Requirements</span></span>

|<span data-ttu-id="8ab66-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-442">Requirement</span></span>| <span data-ttu-id="8ab66-443">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-445">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-445">1.0</span></span>|
|[<span data-ttu-id="8ab66-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-447">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-448">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-449">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8ab66-450">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-450">normalizedSubject: String</span></span>

<span data-ttu-id="8ab66-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8ab66-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="8ab66-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-455">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-455">Type</span></span>

*   <span data-ttu-id="8ab66-456">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-457">Requirements</span></span>

|<span data-ttu-id="8ab66-458">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-458">Requirement</span></span>| <span data-ttu-id="8ab66-459">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-460">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-461">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-461">1.0</span></span>|
|[<span data-ttu-id="8ab66-462">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-463">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-464">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-465">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-466">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="8ab66-467">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-468">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="8ab66-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8ab66-469">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-470">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-470">Read mode</span></span>

<span data-ttu-id="8ab66-471">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="8ab66-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="8ab66-472">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-473">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-474">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-474">Compose mode</span></span>

<span data-ttu-id="8ab66-475">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8ab66-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="8ab66-476">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-477">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ab66-478">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ab66-479">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="8ab66-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-480">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-480">Type</span></span>

*   <span data-ttu-id="8ab66-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-482">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-482">Requirements</span></span>

|<span data-ttu-id="8ab66-483">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-483">Requirement</span></span>| <span data-ttu-id="8ab66-484">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-485">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-486">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-486">1.0</span></span>|
|[<span data-ttu-id="8ab66-487">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-488">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-489">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-490">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="8ab66-491">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-p128">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-494">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-494">Type</span></span>

*   [<span data-ttu-id="8ab66-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ab66-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-496">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-496">Requirements</span></span>

|<span data-ttu-id="8ab66-497">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-497">Requirement</span></span>| <span data-ttu-id="8ab66-498">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-499">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-500">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-500">1.0</span></span>|
|[<span data-ttu-id="8ab66-501">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-502">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-503">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-504">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-505">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="8ab66-506">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-507">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="8ab66-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8ab66-508">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-509">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-509">Read mode</span></span>

<span data-ttu-id="8ab66-510">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="8ab66-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="8ab66-511">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-512">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-513">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-513">Compose mode</span></span>

<span data-ttu-id="8ab66-514">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8ab66-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="8ab66-515">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-516">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ab66-517">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ab66-518">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="8ab66-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-519">Type</span></span>

*   <span data-ttu-id="8ab66-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-521">Requirements</span></span>

|<span data-ttu-id="8ab66-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-522">Requirement</span></span>| <span data-ttu-id="8ab66-523">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-525">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-525">1.0</span></span>|
|[<span data-ttu-id="8ab66-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-527">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-528">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-529">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="8ab66-530">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-p132">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8ab66-p133">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-535">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ab66-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-536">Type</span></span>

*   [<span data-ttu-id="8ab66-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ab66-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="8ab66-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-538">Requirements</span></span>

|<span data-ttu-id="8ab66-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-539">Requirement</span></span>| <span data-ttu-id="8ab66-540">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-542">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-542">1.0</span></span>|
|[<span data-ttu-id="8ab66-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-544">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-546">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="8ab66-548">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-549">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="8ab66-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8ab66-p134">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-552">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-552">Read mode</span></span>

<span data-ttu-id="8ab66-553">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-554">Compose mode</span></span>

<span data-ttu-id="8ab66-555">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8ab66-556">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8ab66-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="8ab66-557">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8ab66-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-558">Type</span></span>

*   <span data-ttu-id="8ab66-559">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-560">Requirements</span></span>

|<span data-ttu-id="8ab66-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-561">Requirement</span></span>| <span data-ttu-id="8ab66-562">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-564">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-564">1.0</span></span>|
|[<span data-ttu-id="8ab66-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-566">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="8ab66-569">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8ab66-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="8ab66-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-572">Read mode</span></span>

<span data-ttu-id="8ab66-p136">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-575">Compose mode</span></span>

<span data-ttu-id="8ab66-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="8ab66-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-577">Type</span></span>

*   <span data-ttu-id="8ab66-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-579">Requirements</span></span>

|<span data-ttu-id="8ab66-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-580">Requirement</span></span>| <span data-ttu-id="8ab66-581">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-583">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-583">1.0</span></span>|
|[<span data-ttu-id="8ab66-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-585">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-586">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-587">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="8ab66-588">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="8ab66-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8ab66-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ab66-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8ab66-591">Read mode</span></span>

<span data-ttu-id="8ab66-592">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="8ab66-593">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-594">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ab66-595">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8ab66-595">Compose mode</span></span>

<span data-ttu-id="8ab66-596">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="8ab66-597">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ab66-598">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="8ab66-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ab66-599">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="8ab66-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ab66-600">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="8ab66-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ab66-601">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-601">Type</span></span>

*   <span data-ttu-id="8ab66-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-603">Requirements</span></span>

|<span data-ttu-id="8ab66-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-604">Requirement</span></span>| <span data-ttu-id="8ab66-605">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-607">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-607">1.0</span></span>|
|[<span data-ttu-id="8ab66-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-609">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-610">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-611">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8ab66-612">Métodos</span><span class="sxs-lookup"><span data-stu-id="8ab66-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8ab66-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ab66-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8ab66-614">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8ab66-615">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="8ab66-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8ab66-616">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8ab66-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-617">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-617">Parameters</span></span>

|<span data-ttu-id="8ab66-618">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-618">Name</span></span>| <span data-ttu-id="8ab66-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-619">Type</span></span>| <span data-ttu-id="8ab66-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-620">Attributes</span></span>| <span data-ttu-id="8ab66-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8ab66-622">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-622">String</span></span>||<span data-ttu-id="8ab66-p140">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8ab66-625">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-625">String</span></span>||<span data-ttu-id="8ab66-p141">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8ab66-628">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-628">Object</span></span>| <span data-ttu-id="8ab66-629">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-629">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-630">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ab66-631">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-631">Object</span></span>| <span data-ttu-id="8ab66-632">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-632">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-633">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ab66-634">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-634">function</span></span>| <span data-ttu-id="8ab66-635">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-635">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-636">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ab66-637">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8ab66-638">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8ab66-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ab66-639">Erros</span><span class="sxs-lookup"><span data-stu-id="8ab66-639">Errors</span></span>

| <span data-ttu-id="8ab66-640">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8ab66-640">Error code</span></span> | <span data-ttu-id="8ab66-641">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8ab66-642">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8ab66-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8ab66-643">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8ab66-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8ab66-644">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8ab66-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-645">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-645">Requirements</span></span>

|<span data-ttu-id="8ab66-646">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-646">Requirement</span></span>| <span data-ttu-id="8ab66-647">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-648">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-649">1.1</span><span class="sxs-lookup"><span data-stu-id="8ab66-649">1.1</span></span>|
|[<span data-ttu-id="8ab66-650">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ab66-652">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-653">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-654">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8ab66-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ab66-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8ab66-656">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8ab66-p142">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8ab66-660">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8ab66-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8ab66-661">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-662">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-662">Parameters</span></span>

|<span data-ttu-id="8ab66-663">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-663">Name</span></span>| <span data-ttu-id="8ab66-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-664">Type</span></span>| <span data-ttu-id="8ab66-665">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-665">Attributes</span></span>| <span data-ttu-id="8ab66-666">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8ab66-667">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-667">String</span></span>||<span data-ttu-id="8ab66-p143">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8ab66-670">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-670">String</span></span>||<span data-ttu-id="8ab66-671">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-671">The subject of the item to be attached.</span></span> <span data-ttu-id="8ab66-672">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8ab66-673">Object</span><span class="sxs-lookup"><span data-stu-id="8ab66-673">Object</span></span>| <span data-ttu-id="8ab66-674">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-674">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-675">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ab66-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-676">Object</span></span>| <span data-ttu-id="8ab66-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-677">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-678">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ab66-679">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-679">function</span></span>| <span data-ttu-id="8ab66-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-680">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-681">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ab66-682">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8ab66-683">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8ab66-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ab66-684">Erros</span><span class="sxs-lookup"><span data-stu-id="8ab66-684">Errors</span></span>

| <span data-ttu-id="8ab66-685">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8ab66-685">Error code</span></span> | <span data-ttu-id="8ab66-686">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8ab66-687">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8ab66-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-688">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-688">Requirements</span></span>

|<span data-ttu-id="8ab66-689">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-689">Requirement</span></span>| <span data-ttu-id="8ab66-690">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-691">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-692">1.1</span><span class="sxs-lookup"><span data-stu-id="8ab66-692">1.1</span></span>|
|[<span data-ttu-id="8ab66-693">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ab66-695">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-696">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-697">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-697">Example</span></span>

<span data-ttu-id="8ab66-698">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8ab66-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8ab66-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8ab66-700">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-701">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ab66-702">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8ab66-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8ab66-703">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8ab66-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8ab66-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-707">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-707">Parameters</span></span>

|<span data-ttu-id="8ab66-708">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-708">Name</span></span>| <span data-ttu-id="8ab66-709">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-709">Type</span></span>| <span data-ttu-id="8ab66-710">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8ab66-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8ab66-711">String &#124; Object</span></span>| |<span data-ttu-id="8ab66-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8ab66-714">**OU**</span><span class="sxs-lookup"><span data-stu-id="8ab66-714">**OR**</span></span><br/><span data-ttu-id="8ab66-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8ab66-717">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-717">String</span></span> | <span data-ttu-id="8ab66-718">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-718">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8ab66-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8ab66-722">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-722">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-723">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8ab66-724">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-724">String</span></span> | | <span data-ttu-id="8ab66-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8ab66-727">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-727">String</span></span> | | <span data-ttu-id="8ab66-728">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8ab66-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8ab66-729">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-729">String</span></span> | | <span data-ttu-id="8ab66-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8ab66-732">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-732">String</span></span> | | <span data-ttu-id="8ab66-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8ab66-736">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-736">function</span></span> | <span data-ttu-id="8ab66-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-737">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-738">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-739">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-739">Requirements</span></span>

|<span data-ttu-id="8ab66-740">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-740">Requirement</span></span>| <span data-ttu-id="8ab66-741">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-742">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-743">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-743">1.0</span></span>|
|[<span data-ttu-id="8ab66-744">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-745">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-746">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-747">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8ab66-748">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8ab66-748">Examples</span></span>

<span data-ttu-id="8ab66-749">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8ab66-750">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8ab66-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8ab66-751">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8ab66-752">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8ab66-753">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8ab66-754">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8ab66-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8ab66-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8ab66-756">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-757">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ab66-758">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8ab66-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8ab66-759">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8ab66-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8ab66-p152">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-763">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-763">Parameters</span></span>

|<span data-ttu-id="8ab66-764">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-764">Name</span></span>| <span data-ttu-id="8ab66-765">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-765">Type</span></span>| <span data-ttu-id="8ab66-766">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8ab66-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8ab66-767">String &#124; Object</span></span>| | <span data-ttu-id="8ab66-p153">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8ab66-770">**OU**</span><span class="sxs-lookup"><span data-stu-id="8ab66-770">**OR**</span></span><br/><span data-ttu-id="8ab66-p154">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8ab66-773">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-773">String</span></span> | <span data-ttu-id="8ab66-774">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-774">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-p155">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8ab66-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8ab66-778">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-778">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-779">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8ab66-780">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-780">String</span></span> | | <span data-ttu-id="8ab66-p156">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8ab66-783">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-783">String</span></span> | | <span data-ttu-id="8ab66-784">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8ab66-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8ab66-785">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-785">String</span></span> | | <span data-ttu-id="8ab66-p157">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8ab66-788">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8ab66-788">String</span></span> | | <span data-ttu-id="8ab66-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8ab66-792">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-792">function</span></span> | <span data-ttu-id="8ab66-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-793">&lt;optional&gt;</span></span> | <span data-ttu-id="8ab66-794">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-795">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-795">Requirements</span></span>

|<span data-ttu-id="8ab66-796">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-796">Requirement</span></span>| <span data-ttu-id="8ab66-797">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-798">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-799">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-799">1.0</span></span>|
|[<span data-ttu-id="8ab66-800">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-801">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-802">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-803">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8ab66-804">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8ab66-804">Examples</span></span>

<span data-ttu-id="8ab66-805">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8ab66-806">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8ab66-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8ab66-807">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8ab66-808">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8ab66-809">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8ab66-810">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="8ab66-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="8ab66-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="8ab66-812">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-813">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-814">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-814">Requirements</span></span>

|<span data-ttu-id="8ab66-815">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-815">Requirement</span></span>| <span data-ttu-id="8ab66-816">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-817">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-818">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-818">1.0</span></span>|
|[<span data-ttu-id="8ab66-819">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-820">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-821">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-822">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-823">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-823">Returns:</span></span>

<span data-ttu-id="8ab66-824">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8ab66-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="8ab66-825">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-825">Example</span></span>

<span data-ttu-id="8ab66-826">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="8ab66-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="8ab66-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="8ab66-828">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-829">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-830">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-830">Parameters</span></span>

|<span data-ttu-id="8ab66-831">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-831">Name</span></span>| <span data-ttu-id="8ab66-832">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-832">Type</span></span>| <span data-ttu-id="8ab66-833">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8ab66-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8ab66-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="8ab66-835">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="8ab66-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ab66-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-836">Requirements</span></span>

|<span data-ttu-id="8ab66-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-837">Requirement</span></span>| <span data-ttu-id="8ab66-838">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-840">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-840">1.0</span></span>|
|[<span data-ttu-id="8ab66-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-842">Restrito</span><span class="sxs-lookup"><span data-stu-id="8ab66-842">Restricted</span></span>|
|[<span data-ttu-id="8ab66-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-844">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-845">Returns:</span></span>

<span data-ttu-id="8ab66-846">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8ab66-847">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8ab66-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8ab66-848">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8ab66-849">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8ab66-850">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8ab66-850">Value of `entityType`</span></span> | <span data-ttu-id="8ab66-851">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="8ab66-851">Type of objects in returned array</span></span> | <span data-ttu-id="8ab66-852">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="8ab66-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8ab66-853">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-853">String</span></span> | <span data-ttu-id="8ab66-854">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8ab66-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8ab66-855">Contato</span><span class="sxs-lookup"><span data-stu-id="8ab66-855">Contact</span></span> | <span data-ttu-id="8ab66-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ab66-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8ab66-857">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-857">String</span></span> | <span data-ttu-id="8ab66-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ab66-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8ab66-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8ab66-859">MeetingSuggestion</span></span> | <span data-ttu-id="8ab66-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ab66-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8ab66-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8ab66-861">PhoneNumber</span></span> | <span data-ttu-id="8ab66-862">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8ab66-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8ab66-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8ab66-863">TaskSuggestion</span></span> | <span data-ttu-id="8ab66-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ab66-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8ab66-865">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-865">String</span></span> | <span data-ttu-id="8ab66-866">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8ab66-866">**Restricted**</span></span> |

<span data-ttu-id="8ab66-867">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="8ab66-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="8ab66-868">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-868">Example</span></span>

<span data-ttu-id="8ab66-869">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8ab66-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="8ab66-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="8ab66-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="8ab66-871">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-872">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ab66-873">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-874">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-874">Parameters</span></span>

|<span data-ttu-id="8ab66-875">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-875">Name</span></span>| <span data-ttu-id="8ab66-876">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-876">Type</span></span>| <span data-ttu-id="8ab66-877">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8ab66-878">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-878">String</span></span>|<span data-ttu-id="8ab66-879">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8ab66-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ab66-880">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-880">Requirements</span></span>

|<span data-ttu-id="8ab66-881">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-881">Requirement</span></span>| <span data-ttu-id="8ab66-882">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-883">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-884">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-884">1.0</span></span>|
|[<span data-ttu-id="8ab66-885">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-886">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-887">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-888">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-889">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-889">Returns:</span></span>

<span data-ttu-id="8ab66-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8ab66-892">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="8ab66-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="8ab66-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8ab66-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8ab66-894">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-895">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ab66-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8ab66-899">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="8ab66-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8ab66-900">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="8ab66-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ab66-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-903">Requirements</span></span>

|<span data-ttu-id="8ab66-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-904">Requirement</span></span>| <span data-ttu-id="8ab66-905">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-906">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-907">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-907">1.0</span></span>|
|[<span data-ttu-id="8ab66-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-909">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-910">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-911">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-912">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-912">Returns:</span></span>

<span data-ttu-id="8ab66-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="8ab66-915">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="8ab66-916">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-916">Example</span></span>

<span data-ttu-id="8ab66-917">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="8ab66-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8ab66-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8ab66-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8ab66-919">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-920">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="8ab66-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ab66-921">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8ab66-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-924">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-924">Parameters</span></span>

|<span data-ttu-id="8ab66-925">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-925">Name</span></span>| <span data-ttu-id="8ab66-926">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-926">Type</span></span>| <span data-ttu-id="8ab66-927">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8ab66-928">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-928">String</span></span>|<span data-ttu-id="8ab66-929">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="8ab66-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ab66-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-930">Requirements</span></span>

|<span data-ttu-id="8ab66-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-931">Requirement</span></span>| <span data-ttu-id="8ab66-932">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-934">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-934">1.0</span></span>|
|[<span data-ttu-id="8ab66-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-936">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-938">Read</span><span class="sxs-lookup"><span data-stu-id="8ab66-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-939">Returns:</span></span>

<span data-ttu-id="8ab66-940">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="8ab66-941">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8ab66-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="8ab66-942">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8ab66-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8ab66-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8ab66-944">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8ab66-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="8ab66-947">No Outlook na Web, o método retorna a cadeia de caracteres "NULL" se nenhum texto está selecionado, mas o cursor está no corpo.</span><span class="sxs-lookup"><span data-stu-id="8ab66-947">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="8ab66-948">Para verificar essa situação, inclua um código semelhante ao seguinte:</span><span class="sxs-lookup"><span data-stu-id="8ab66-948">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="8ab66-949">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-949">Parameters</span></span>

|<span data-ttu-id="8ab66-950">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-950">Name</span></span>| <span data-ttu-id="8ab66-951">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-951">Type</span></span>| <span data-ttu-id="8ab66-952">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-952">Attributes</span></span>| <span data-ttu-id="8ab66-953">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-953">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8ab66-954">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8ab66-954">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8ab66-p167">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8ab66-958">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-958">Object</span></span>| <span data-ttu-id="8ab66-959">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-959">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-960">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ab66-961">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-961">Object</span></span>| <span data-ttu-id="8ab66-962">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-962">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-963">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ab66-964">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-964">function</span></span>||<span data-ttu-id="8ab66-965">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-965">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8ab66-966">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-966">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8ab66-967">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-967">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ab66-968">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-968">Requirements</span></span>

|<span data-ttu-id="8ab66-969">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-969">Requirement</span></span>| <span data-ttu-id="8ab66-970">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-971">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-972">1.2</span><span class="sxs-lookup"><span data-stu-id="8ab66-972">1.2</span></span>|
|[<span data-ttu-id="8ab66-973">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-974">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-975">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-976">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-976">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ab66-977">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8ab66-977">Returns:</span></span>

<span data-ttu-id="8ab66-978">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-978">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="8ab66-979">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="8ab66-979">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8ab66-980">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-980">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8ab66-981">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8ab66-981">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8ab66-982">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-982">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8ab66-p169">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-986">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-986">Parameters</span></span>

|<span data-ttu-id="8ab66-987">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-987">Name</span></span>| <span data-ttu-id="8ab66-988">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-988">Type</span></span>| <span data-ttu-id="8ab66-989">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-989">Attributes</span></span>| <span data-ttu-id="8ab66-990">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-990">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8ab66-991">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-991">function</span></span>||<span data-ttu-id="8ab66-992">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8ab66-993">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-993">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8ab66-994">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="8ab66-994">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8ab66-995">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-995">Object</span></span>| <span data-ttu-id="8ab66-996">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-996">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-997">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-997">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8ab66-998">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-998">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ab66-999">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-999">Requirements</span></span>

|<span data-ttu-id="8ab66-1000">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-1000">Requirement</span></span>| <span data-ttu-id="8ab66-1001">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-1002">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="8ab66-1003">1.0</span></span>|
|[<span data-ttu-id="8ab66-1004">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-1005">ReadItem</span></span>|
|[<span data-ttu-id="8ab66-1006">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8ab66-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-1007">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8ab66-1007">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1008">Example</span></span>

<span data-ttu-id="8ab66-p172">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8ab66-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ab66-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8ab66-1013">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1013">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8ab66-1014">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1014">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8ab66-1015">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1015">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8ab66-1016">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1016">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8ab66-1017">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1017">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-1018">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-1018">Parameters</span></span>

|<span data-ttu-id="8ab66-1019">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-1019">Name</span></span>| <span data-ttu-id="8ab66-1020">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1020">Type</span></span>| <span data-ttu-id="8ab66-1021">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-1021">Attributes</span></span>| <span data-ttu-id="8ab66-1022">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-1022">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8ab66-1023">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-1023">String</span></span>||<span data-ttu-id="8ab66-1024">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1024">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8ab66-1025">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-1025">Object</span></span>| <span data-ttu-id="8ab66-1026">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1027">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1027">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ab66-1028">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-1028">Object</span></span>| <span data-ttu-id="8ab66-1029">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1029">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1030">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1030">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ab66-1031">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-1031">function</span></span>| <span data-ttu-id="8ab66-1032">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1032">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1033">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-1033">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ab66-1034">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1034">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ab66-1035">Erros</span><span class="sxs-lookup"><span data-stu-id="8ab66-1035">Errors</span></span>

| <span data-ttu-id="8ab66-1036">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8ab66-1036">Error code</span></span> | <span data-ttu-id="8ab66-1037">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-1037">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8ab66-1038">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1038">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-1039">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-1039">Requirements</span></span>

|<span data-ttu-id="8ab66-1040">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-1040">Requirement</span></span>| <span data-ttu-id="8ab66-1041">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-1042">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-1043">1.1</span><span class="sxs-lookup"><span data-stu-id="8ab66-1043">1.1</span></span>|
|[<span data-ttu-id="8ab66-1044">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1044">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-1045">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-1045">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ab66-1046">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-1046">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-1047">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-1047">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-1048">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1048">Example</span></span>

<span data-ttu-id="8ab66-1049">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1049">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8ab66-1050">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8ab66-1050">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8ab66-1051">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1051">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8ab66-p174">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ab66-1055">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="8ab66-1055">Parameters</span></span>

|<span data-ttu-id="8ab66-1056">Nome</span><span class="sxs-lookup"><span data-stu-id="8ab66-1056">Name</span></span>| <span data-ttu-id="8ab66-1057">Tipo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1057">Type</span></span>| <span data-ttu-id="8ab66-1058">Atributos</span><span class="sxs-lookup"><span data-stu-id="8ab66-1058">Attributes</span></span>| <span data-ttu-id="8ab66-1059">Descrição</span><span class="sxs-lookup"><span data-stu-id="8ab66-1059">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8ab66-1060">String</span><span class="sxs-lookup"><span data-stu-id="8ab66-1060">String</span></span>||<span data-ttu-id="8ab66-p175">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8ab66-1064">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-1064">Object</span></span>| <span data-ttu-id="8ab66-1065">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1065">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1066">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1066">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ab66-1067">Objeto</span><span class="sxs-lookup"><span data-stu-id="8ab66-1067">Object</span></span>| <span data-ttu-id="8ab66-1068">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1069">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1069">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8ab66-1070">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8ab66-1070">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8ab66-1071">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ab66-1071">&lt;optional&gt;</span></span>|<span data-ttu-id="8ab66-1072">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1072">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="8ab66-1073">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1073">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8ab66-1074">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1074">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="8ab66-1075">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1075">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8ab66-1076">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="8ab66-1076">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8ab66-1077">function</span><span class="sxs-lookup"><span data-stu-id="8ab66-1077">function</span></span>||<span data-ttu-id="8ab66-1078">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8ab66-1078">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ab66-1079">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8ab66-1079">Requirements</span></span>

|<span data-ttu-id="8ab66-1080">Requisito</span><span class="sxs-lookup"><span data-stu-id="8ab66-1080">Requirement</span></span>| <span data-ttu-id="8ab66-1081">Valor</span><span class="sxs-lookup"><span data-stu-id="8ab66-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ab66-1082">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8ab66-1082">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ab66-1083">1.2</span><span class="sxs-lookup"><span data-stu-id="8ab66-1083">1.2</span></span>|
|[<span data-ttu-id="8ab66-1084">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1084">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ab66-1085">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ab66-1085">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ab66-1086">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8ab66-1086">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ab66-1087">Escrever</span><span class="sxs-lookup"><span data-stu-id="8ab66-1087">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ab66-1088">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8ab66-1088">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
