---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,2
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: ab8c55d2f91b250b419c7c9c71fc044b6fa68279
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629206"
---
# <a name="item"></a><span data-ttu-id="7c6be-102">item</span><span class="sxs-lookup"><span data-stu-id="7c6be-102">item</span></span>

### <span data-ttu-id="7c6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="7c6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="7c6be-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="7c6be-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-107">Requirements</span></span>

|<span data-ttu-id="7c6be-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-108">Requirement</span></span>| <span data-ttu-id="7c6be-109">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-111">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-111">1.0</span></span>|
|[<span data-ttu-id="7c6be-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="7c6be-113">Restricted</span></span>|
|[<span data-ttu-id="7c6be-114">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-115">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7c6be-116">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="7c6be-116">Members and methods</span></span>

| <span data-ttu-id="7c6be-117">Membro	</span><span class="sxs-lookup"><span data-stu-id="7c6be-117">Member</span></span> | <span data-ttu-id="7c6be-118">Tipo	</span><span class="sxs-lookup"><span data-stu-id="7c6be-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7c6be-119">attachments</span><span class="sxs-lookup"><span data-stu-id="7c6be-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="7c6be-120">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-120">Member</span></span> |
| [<span data-ttu-id="7c6be-121">bcc</span><span class="sxs-lookup"><span data-stu-id="7c6be-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="7c6be-122">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-122">Member</span></span> |
| [<span data-ttu-id="7c6be-123">body</span><span class="sxs-lookup"><span data-stu-id="7c6be-123">body</span></span>](#body-body) | <span data-ttu-id="7c6be-124">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-124">Member</span></span> |
| [<span data-ttu-id="7c6be-125">cc</span><span class="sxs-lookup"><span data-stu-id="7c6be-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7c6be-126">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-126">Member</span></span> |
| [<span data-ttu-id="7c6be-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="7c6be-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="7c6be-128">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-128">Member</span></span> |
| [<span data-ttu-id="7c6be-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="7c6be-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="7c6be-130">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-130">Member</span></span> |
| [<span data-ttu-id="7c6be-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="7c6be-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="7c6be-132">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-132">Member</span></span> |
| [<span data-ttu-id="7c6be-133">end</span><span class="sxs-lookup"><span data-stu-id="7c6be-133">end</span></span>](#end-datetime) | <span data-ttu-id="7c6be-134">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-134">Member</span></span> |
| [<span data-ttu-id="7c6be-135">from</span><span class="sxs-lookup"><span data-stu-id="7c6be-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="7c6be-136">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-136">Member</span></span> |
| [<span data-ttu-id="7c6be-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="7c6be-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="7c6be-138">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-138">Member</span></span> |
| [<span data-ttu-id="7c6be-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="7c6be-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="7c6be-140">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-140">Member</span></span> |
| [<span data-ttu-id="7c6be-141">itemId</span><span class="sxs-lookup"><span data-stu-id="7c6be-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="7c6be-142">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-142">Member</span></span> |
| [<span data-ttu-id="7c6be-143">itemType</span><span class="sxs-lookup"><span data-stu-id="7c6be-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="7c6be-144">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-144">Member</span></span> |
| [<span data-ttu-id="7c6be-145">location</span><span class="sxs-lookup"><span data-stu-id="7c6be-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="7c6be-146">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-146">Member</span></span> |
| [<span data-ttu-id="7c6be-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="7c6be-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="7c6be-148">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-148">Member</span></span> |
| [<span data-ttu-id="7c6be-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="7c6be-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7c6be-150">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-150">Member</span></span> |
| [<span data-ttu-id="7c6be-151">organizer</span><span class="sxs-lookup"><span data-stu-id="7c6be-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="7c6be-152">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-152">Member</span></span> |
| [<span data-ttu-id="7c6be-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="7c6be-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7c6be-154">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-154">Member</span></span> |
| [<span data-ttu-id="7c6be-155">sender</span><span class="sxs-lookup"><span data-stu-id="7c6be-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="7c6be-156">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-156">Member</span></span> |
| [<span data-ttu-id="7c6be-157">start</span><span class="sxs-lookup"><span data-stu-id="7c6be-157">start</span></span>](#start-datetime) | <span data-ttu-id="7c6be-158">Member</span><span class="sxs-lookup"><span data-stu-id="7c6be-158">Member</span></span> |
| [<span data-ttu-id="7c6be-159">subject</span><span class="sxs-lookup"><span data-stu-id="7c6be-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="7c6be-160">Membro</span><span class="sxs-lookup"><span data-stu-id="7c6be-160">Member</span></span> |
| [<span data-ttu-id="7c6be-161">to</span><span class="sxs-lookup"><span data-stu-id="7c6be-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7c6be-162">Membro</span><span class="sxs-lookup"><span data-stu-id="7c6be-162">Member</span></span> |
| [<span data-ttu-id="7c6be-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="7c6be-164">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-164">Method</span></span> |
| [<span data-ttu-id="7c6be-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="7c6be-166">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-166">Method</span></span> |
| [<span data-ttu-id="7c6be-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="7c6be-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="7c6be-168">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-168">Method</span></span> |
| [<span data-ttu-id="7c6be-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="7c6be-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="7c6be-170">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-170">Method</span></span> |
| [<span data-ttu-id="7c6be-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="7c6be-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="7c6be-172">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-172">Method</span></span> |
| [<span data-ttu-id="7c6be-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="7c6be-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7c6be-174">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-174">Method</span></span> |
| [<span data-ttu-id="7c6be-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="7c6be-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7c6be-176">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-176">Method</span></span> |
| [<span data-ttu-id="7c6be-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="7c6be-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="7c6be-178">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-178">Method</span></span> |
| [<span data-ttu-id="7c6be-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="7c6be-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="7c6be-180">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-180">Method</span></span> |
| [<span data-ttu-id="7c6be-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="7c6be-182">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-182">Method</span></span> |
| [<span data-ttu-id="7c6be-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="7c6be-184">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-184">Method</span></span> |
| [<span data-ttu-id="7c6be-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="7c6be-186">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-186">Method</span></span> |
| [<span data-ttu-id="7c6be-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7c6be-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="7c6be-188">Método</span><span class="sxs-lookup"><span data-stu-id="7c6be-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="7c6be-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-189">Example</span></span>

<span data-ttu-id="7c6be-190">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="7c6be-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="7c6be-191">Members</span><span class="sxs-lookup"><span data-stu-id="7c6be-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="7c6be-192">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="7c6be-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="7c6be-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-195">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="7c6be-196">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="7c6be-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-197">Type</span></span>

*   <span data-ttu-id="7c6be-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="7c6be-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-199">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-199">Requirements</span></span>

|<span data-ttu-id="7c6be-200">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-200">Requirement</span></span>| <span data-ttu-id="7c6be-201">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-202">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-203">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-203">1.0</span></span>|
|[<span data-ttu-id="7c6be-204">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-205">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-206">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-207">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-208">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-208">Example</span></span>

<span data-ttu-id="7c6be-209">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="7c6be-210">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-211">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="7c6be-212">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="7c6be-212">Compose mode only.</span></span>

<span data-ttu-id="7c6be-213">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-214">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7c6be-215">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="7c6be-216">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="7c6be-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-217">Type</span></span>

*   [<span data-ttu-id="7c6be-218">Destinatários</span><span class="sxs-lookup"><span data-stu-id="7c6be-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-219">Requirements</span></span>

|<span data-ttu-id="7c6be-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-220">Requirement</span></span>| <span data-ttu-id="7c6be-221">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-223">1.1</span><span class="sxs-lookup"><span data-stu-id="7c6be-223">1.1</span></span>|
|[<span data-ttu-id="7c6be-224">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-225">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-227">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-228">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="7c6be-229">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-230">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-231">Type</span></span>

*   [<span data-ttu-id="7c6be-232">Body</span><span class="sxs-lookup"><span data-stu-id="7c6be-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-233">Requirements</span></span>

|<span data-ttu-id="7c6be-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-234">Requirement</span></span>| <span data-ttu-id="7c6be-235">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-237">1.1</span><span class="sxs-lookup"><span data-stu-id="7c6be-237">1.1</span></span>|
|[<span data-ttu-id="7c6be-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-239">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-240">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-241">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-242">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-242">Example</span></span>

<span data-ttu-id="7c6be-243">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="7c6be-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="7c6be-244">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="7c6be-245">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-246">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="7c6be-247">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-248">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-248">Read mode</span></span>

<span data-ttu-id="7c6be-249">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="7c6be-250">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-251">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-252">Compose mode</span></span>

<span data-ttu-id="7c6be-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="7c6be-254">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-255">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7c6be-256">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="7c6be-257">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="7c6be-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-258">Type</span></span>

*   <span data-ttu-id="7c6be-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-260">Requirements</span></span>

|<span data-ttu-id="7c6be-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-261">Requirement</span></span>| <span data-ttu-id="7c6be-262">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-264">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-264">1.0</span></span>|
|[<span data-ttu-id="7c6be-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-266">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="7c6be-269">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="7c6be-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="7c6be-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="7c6be-p110">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="7c6be-p111">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-275">Type</span></span>

*   <span data-ttu-id="7c6be-276">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-277">Requirements</span></span>

|<span data-ttu-id="7c6be-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-278">Requirement</span></span>| <span data-ttu-id="7c6be-279">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-281">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-281">1.0</span></span>|
|[<span data-ttu-id="7c6be-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-283">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="7c6be-287">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="7c6be-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="7c6be-p112">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-290">Type</span></span>

*   <span data-ttu-id="7c6be-291">Data</span><span class="sxs-lookup"><span data-stu-id="7c6be-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-292">Requirements</span></span>

|<span data-ttu-id="7c6be-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-293">Requirement</span></span>| <span data-ttu-id="7c6be-294">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-296">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-296">1.0</span></span>|
|[<span data-ttu-id="7c6be-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-298">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-300">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="7c6be-302">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="7c6be-302">dateTimeModified: Date</span></span>

<span data-ttu-id="7c6be-p113">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-305">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-306">Type</span></span>

*   <span data-ttu-id="7c6be-307">Data</span><span class="sxs-lookup"><span data-stu-id="7c6be-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-308">Requirements</span></span>

|<span data-ttu-id="7c6be-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-309">Requirement</span></span>| <span data-ttu-id="7c6be-310">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-312">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-312">1.0</span></span>|
|[<span data-ttu-id="7c6be-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-314">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-316">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="7c6be-318">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="7c6be-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="7c6be-p114">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-322">Read mode</span></span>

<span data-ttu-id="7c6be-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-324">Compose mode</span></span>

<span data-ttu-id="7c6be-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="7c6be-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="7c6be-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7c6be-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7c6be-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-328">Type</span></span>

*   <span data-ttu-id="7c6be-329">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-330">Requirements</span></span>

|<span data-ttu-id="7c6be-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-331">Requirement</span></span>| <span data-ttu-id="7c6be-332">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-334">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-334">1.0</span></span>|
|[<span data-ttu-id="7c6be-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-336">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="7c6be-339">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-p115">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="7c6be-p116">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-345">Type</span></span>

*   [<span data-ttu-id="7c6be-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7c6be-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-347">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-347">Requirements</span></span>

|<span data-ttu-id="7c6be-348">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-348">Requirement</span></span>| <span data-ttu-id="7c6be-349">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-350">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-351">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-351">1.0</span></span>|
|[<span data-ttu-id="7c6be-352">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-353">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-354">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-355">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-356">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="7c6be-357">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-357">internetMessageId: String</span></span>

<span data-ttu-id="7c6be-p117">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-360">Type</span></span>

*   <span data-ttu-id="7c6be-361">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-362">Requirements</span></span>

|<span data-ttu-id="7c6be-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-363">Requirement</span></span>| <span data-ttu-id="7c6be-364">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-366">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-366">1.0</span></span>|
|[<span data-ttu-id="7c6be-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-368">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-370">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="7c6be-372">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="7c6be-372">itemClass: String</span></span>

<span data-ttu-id="7c6be-p118">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="7c6be-p119">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="7c6be-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-377">Type</span></span> | <span data-ttu-id="7c6be-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-378">Description</span></span> | <span data-ttu-id="7c6be-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="7c6be-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="7c6be-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="7c6be-380">Appointment items</span></span> | <span data-ttu-id="7c6be-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="7c6be-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="7c6be-382">Message items</span></span> | <span data-ttu-id="7c6be-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="7c6be-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="7c6be-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-385">Type</span></span>

*   <span data-ttu-id="7c6be-386">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-387">Requirements</span></span>

|<span data-ttu-id="7c6be-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-388">Requirement</span></span>| <span data-ttu-id="7c6be-389">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-391">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-391">1.0</span></span>|
|[<span data-ttu-id="7c6be-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-393">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-395">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="7c6be-397">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-397">(nullable) itemId: String</span></span>

<span data-ttu-id="7c6be-p120">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p120">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-400">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="7c6be-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="7c6be-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7c6be-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="7c6be-402">Antes de fazer chamadas da API REST usando esse valor, ele deve ser `Office.context.mailbox.convertToRestId`convertido usando o, que está disponível a partir do conjunto de requisitos 1,3.</span><span class="sxs-lookup"><span data-stu-id="7c6be-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="7c6be-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="7c6be-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-404">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-404">Type</span></span>

*   <span data-ttu-id="7c6be-405">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-406">Requirements</span></span>

|<span data-ttu-id="7c6be-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-407">Requirement</span></span>| <span data-ttu-id="7c6be-408">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-410">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-410">1.0</span></span>|
|[<span data-ttu-id="7c6be-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-412">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-414">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-415">Example</span></span>

<span data-ttu-id="7c6be-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="7c6be-418">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-419">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="7c6be-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="7c6be-420">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-421">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-421">Type</span></span>

*   [<span data-ttu-id="7c6be-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="7c6be-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-423">Requirements</span></span>

|<span data-ttu-id="7c6be-424">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-424">Requirement</span></span>| <span data-ttu-id="7c6be-425">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-426">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-427">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-427">1.0</span></span>|
|[<span data-ttu-id="7c6be-428">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-429">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-430">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-431">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-432">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="7c6be-433">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-434">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-435">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-435">Read mode</span></span>

<span data-ttu-id="7c6be-436">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-437">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-437">Compose mode</span></span>

<span data-ttu-id="7c6be-438">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-439">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-439">Type</span></span>

*   <span data-ttu-id="7c6be-440">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-441">Requirements</span></span>

|<span data-ttu-id="7c6be-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-442">Requirement</span></span>| <span data-ttu-id="7c6be-443">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-445">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-445">1.0</span></span>|
|[<span data-ttu-id="7c6be-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-447">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-448">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-449">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="7c6be-450">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-450">normalizedSubject: String</span></span>

<span data-ttu-id="7c6be-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="7c6be-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="7c6be-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-455">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-455">Type</span></span>

*   <span data-ttu-id="7c6be-456">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-457">Requirements</span></span>

|<span data-ttu-id="7c6be-458">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-458">Requirement</span></span>| <span data-ttu-id="7c6be-459">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-460">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-461">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-461">1.0</span></span>|
|[<span data-ttu-id="7c6be-462">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-463">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-464">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-465">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-466">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="7c6be-467">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-468">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="7c6be-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="7c6be-469">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-470">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-470">Read mode</span></span>

<span data-ttu-id="7c6be-471">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="7c6be-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="7c6be-472">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-473">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-474">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-474">Compose mode</span></span>

<span data-ttu-id="7c6be-475">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="7c6be-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="7c6be-476">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-477">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7c6be-478">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="7c6be-479">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="7c6be-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-480">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-480">Type</span></span>

*   <span data-ttu-id="7c6be-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-482">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-482">Requirements</span></span>

|<span data-ttu-id="7c6be-483">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-483">Requirement</span></span>| <span data-ttu-id="7c6be-484">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-485">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-486">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-486">1.0</span></span>|
|[<span data-ttu-id="7c6be-487">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-488">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-489">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-490">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="7c6be-491">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-p128">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-494">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-494">Type</span></span>

*   [<span data-ttu-id="7c6be-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7c6be-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-496">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-496">Requirements</span></span>

|<span data-ttu-id="7c6be-497">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-497">Requirement</span></span>| <span data-ttu-id="7c6be-498">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-499">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-500">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-500">1.0</span></span>|
|[<span data-ttu-id="7c6be-501">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-502">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-503">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-504">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-505">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="7c6be-506">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-507">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="7c6be-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="7c6be-508">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-509">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-509">Read mode</span></span>

<span data-ttu-id="7c6be-510">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="7c6be-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="7c6be-511">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-512">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-513">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-513">Compose mode</span></span>

<span data-ttu-id="7c6be-514">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="7c6be-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="7c6be-515">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-516">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7c6be-517">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="7c6be-518">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="7c6be-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-519">Type</span></span>

*   <span data-ttu-id="7c6be-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-521">Requirements</span></span>

|<span data-ttu-id="7c6be-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-522">Requirement</span></span>| <span data-ttu-id="7c6be-523">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-525">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-525">1.0</span></span>|
|[<span data-ttu-id="7c6be-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-527">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-528">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-529">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="7c6be-530">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-p132">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="7c6be-p133">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-535">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7c6be-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-536">Type</span></span>

*   [<span data-ttu-id="7c6be-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7c6be-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="7c6be-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-538">Requirements</span></span>

|<span data-ttu-id="7c6be-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-539">Requirement</span></span>| <span data-ttu-id="7c6be-540">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-542">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-542">1.0</span></span>|
|[<span data-ttu-id="7c6be-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-544">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-546">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="7c6be-548">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-549">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="7c6be-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="7c6be-p134">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-552">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-552">Read mode</span></span>

<span data-ttu-id="7c6be-553">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-554">Compose mode</span></span>

<span data-ttu-id="7c6be-555">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="7c6be-556">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="7c6be-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="7c6be-557">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7c6be-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-558">Type</span></span>

*   <span data-ttu-id="7c6be-559">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-560">Requirements</span></span>

|<span data-ttu-id="7c6be-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-561">Requirement</span></span>| <span data-ttu-id="7c6be-562">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-564">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-564">1.0</span></span>|
|[<span data-ttu-id="7c6be-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-566">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="7c6be-569">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="7c6be-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="7c6be-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-572">Read mode</span></span>

<span data-ttu-id="7c6be-p136">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-575">Compose mode</span></span>

<span data-ttu-id="7c6be-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="7c6be-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-577">Type</span></span>

*   <span data-ttu-id="7c6be-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-579">Requirements</span></span>

|<span data-ttu-id="7c6be-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-580">Requirement</span></span>| <span data-ttu-id="7c6be-581">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-583">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-583">1.0</span></span>|
|[<span data-ttu-id="7c6be-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-585">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-586">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-587">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="7c6be-588">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="7c6be-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="7c6be-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7c6be-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="7c6be-591">Read mode</span></span>

<span data-ttu-id="7c6be-592">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="7c6be-593">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-594">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="7c6be-595">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="7c6be-595">Compose mode</span></span>

<span data-ttu-id="7c6be-596">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="7c6be-597">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="7c6be-598">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="7c6be-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="7c6be-599">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="7c6be-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="7c6be-600">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="7c6be-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7c6be-601">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-601">Type</span></span>

*   <span data-ttu-id="7c6be-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-603">Requirements</span></span>

|<span data-ttu-id="7c6be-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-604">Requirement</span></span>| <span data-ttu-id="7c6be-605">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-607">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-607">1.0</span></span>|
|[<span data-ttu-id="7c6be-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-609">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-610">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-611">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7c6be-612">Métodos</span><span class="sxs-lookup"><span data-stu-id="7c6be-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="7c6be-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7c6be-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7c6be-614">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="7c6be-615">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="7c6be-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="7c6be-616">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="7c6be-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-617">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-617">Parameters</span></span>

|<span data-ttu-id="7c6be-618">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-618">Name</span></span>| <span data-ttu-id="7c6be-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-619">Type</span></span>| <span data-ttu-id="7c6be-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-620">Attributes</span></span>| <span data-ttu-id="7c6be-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="7c6be-622">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-622">String</span></span>||<span data-ttu-id="7c6be-p140">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7c6be-625">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-625">String</span></span>||<span data-ttu-id="7c6be-p141">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7c6be-628">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-628">Object</span></span>| <span data-ttu-id="7c6be-629">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-629">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-630">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7c6be-631">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-631">Object</span></span>| <span data-ttu-id="7c6be-632">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-632">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-633">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7c6be-634">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-634">function</span></span>| <span data-ttu-id="7c6be-635">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-635">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-636">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7c6be-637">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7c6be-638">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="7c6be-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7c6be-639">Erros</span><span class="sxs-lookup"><span data-stu-id="7c6be-639">Errors</span></span>

| <span data-ttu-id="7c6be-640">Código de erro</span><span class="sxs-lookup"><span data-stu-id="7c6be-640">Error code</span></span> | <span data-ttu-id="7c6be-641">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="7c6be-642">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="7c6be-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="7c6be-643">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="7c6be-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7c6be-644">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="7c6be-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-645">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-645">Requirements</span></span>

|<span data-ttu-id="7c6be-646">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-646">Requirement</span></span>| <span data-ttu-id="7c6be-647">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-648">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-649">1.1</span><span class="sxs-lookup"><span data-stu-id="7c6be-649">1.1</span></span>|
|[<span data-ttu-id="7c6be-650">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="7c6be-652">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-653">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-654">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="7c6be-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7c6be-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7c6be-656">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="7c6be-p142">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="7c6be-660">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="7c6be-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="7c6be-661">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-662">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-662">Parameters</span></span>

|<span data-ttu-id="7c6be-663">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-663">Name</span></span>| <span data-ttu-id="7c6be-664">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-664">Type</span></span>| <span data-ttu-id="7c6be-665">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-665">Attributes</span></span>| <span data-ttu-id="7c6be-666">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="7c6be-667">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-667">String</span></span>||<span data-ttu-id="7c6be-p143">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7c6be-670">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-670">String</span></span>||<span data-ttu-id="7c6be-671">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-671">The subject of the item to be attached.</span></span> <span data-ttu-id="7c6be-672">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7c6be-673">Object</span><span class="sxs-lookup"><span data-stu-id="7c6be-673">Object</span></span>| <span data-ttu-id="7c6be-674">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-674">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-675">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7c6be-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-676">Object</span></span>| <span data-ttu-id="7c6be-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-677">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-678">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7c6be-679">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-679">function</span></span>| <span data-ttu-id="7c6be-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-680">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-681">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7c6be-682">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7c6be-683">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="7c6be-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7c6be-684">Erros</span><span class="sxs-lookup"><span data-stu-id="7c6be-684">Errors</span></span>

| <span data-ttu-id="7c6be-685">Código de erro</span><span class="sxs-lookup"><span data-stu-id="7c6be-685">Error code</span></span> | <span data-ttu-id="7c6be-686">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7c6be-687">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="7c6be-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-688">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-688">Requirements</span></span>

|<span data-ttu-id="7c6be-689">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-689">Requirement</span></span>| <span data-ttu-id="7c6be-690">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-691">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-692">1.1</span><span class="sxs-lookup"><span data-stu-id="7c6be-692">1.1</span></span>|
|[<span data-ttu-id="7c6be-693">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="7c6be-695">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-696">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-697">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-697">Example</span></span>

<span data-ttu-id="7c6be-698">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="7c6be-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7c6be-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="7c6be-700">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-701">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7c6be-702">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="7c6be-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7c6be-703">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="7c6be-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="7c6be-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-707">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-707">Parameters</span></span>

|<span data-ttu-id="7c6be-708">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-708">Name</span></span>| <span data-ttu-id="7c6be-709">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-709">Type</span></span>| <span data-ttu-id="7c6be-710">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7c6be-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7c6be-711">String &#124; Object</span></span>| |<span data-ttu-id="7c6be-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7c6be-714">**OU**</span><span class="sxs-lookup"><span data-stu-id="7c6be-714">**OR**</span></span><br/><span data-ttu-id="7c6be-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7c6be-717">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-717">String</span></span> | <span data-ttu-id="7c6be-718">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-718">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7c6be-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7c6be-722">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-722">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-723">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7c6be-724">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-724">String</span></span> | | <span data-ttu-id="7c6be-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7c6be-727">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-727">String</span></span> | | <span data-ttu-id="7c6be-728">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="7c6be-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7c6be-729">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-729">String</span></span> | | <span data-ttu-id="7c6be-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7c6be-732">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-732">String</span></span> | | <span data-ttu-id="7c6be-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7c6be-736">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-736">function</span></span> | <span data-ttu-id="7c6be-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-737">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-738">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-739">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-739">Requirements</span></span>

|<span data-ttu-id="7c6be-740">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-740">Requirement</span></span>| <span data-ttu-id="7c6be-741">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-742">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-743">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-743">1.0</span></span>|
|[<span data-ttu-id="7c6be-744">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-745">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-746">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-747">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7c6be-748">Exemplos</span><span class="sxs-lookup"><span data-stu-id="7c6be-748">Examples</span></span>

<span data-ttu-id="7c6be-749">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="7c6be-750">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="7c6be-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="7c6be-751">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7c6be-752">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7c6be-753">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7c6be-754">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="7c6be-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7c6be-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="7c6be-756">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-757">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7c6be-758">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="7c6be-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7c6be-759">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="7c6be-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="7c6be-p152">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-763">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-763">Parameters</span></span>

|<span data-ttu-id="7c6be-764">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-764">Name</span></span>| <span data-ttu-id="7c6be-765">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-765">Type</span></span>| <span data-ttu-id="7c6be-766">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7c6be-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7c6be-767">String &#124; Object</span></span>| | <span data-ttu-id="7c6be-p153">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7c6be-770">**OU**</span><span class="sxs-lookup"><span data-stu-id="7c6be-770">**OR**</span></span><br/><span data-ttu-id="7c6be-p154">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7c6be-773">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-773">String</span></span> | <span data-ttu-id="7c6be-774">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-774">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-p155">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7c6be-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7c6be-778">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-778">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-779">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7c6be-780">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-780">String</span></span> | | <span data-ttu-id="7c6be-p156">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7c6be-783">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-783">String</span></span> | | <span data-ttu-id="7c6be-784">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="7c6be-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7c6be-785">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-785">String</span></span> | | <span data-ttu-id="7c6be-p157">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7c6be-788">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c6be-788">String</span></span> | | <span data-ttu-id="7c6be-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7c6be-792">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-792">function</span></span> | <span data-ttu-id="7c6be-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-793">&lt;optional&gt;</span></span> | <span data-ttu-id="7c6be-794">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-795">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-795">Requirements</span></span>

|<span data-ttu-id="7c6be-796">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-796">Requirement</span></span>| <span data-ttu-id="7c6be-797">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-798">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-799">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-799">1.0</span></span>|
|[<span data-ttu-id="7c6be-800">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-801">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-802">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-803">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7c6be-804">Exemplos</span><span class="sxs-lookup"><span data-stu-id="7c6be-804">Examples</span></span>

<span data-ttu-id="7c6be-805">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="7c6be-806">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="7c6be-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="7c6be-807">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7c6be-808">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7c6be-809">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7c6be-810">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="7c6be-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="7c6be-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="7c6be-812">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-813">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-814">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-814">Requirements</span></span>

|<span data-ttu-id="7c6be-815">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-815">Requirement</span></span>| <span data-ttu-id="7c6be-816">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-817">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-818">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-818">1.0</span></span>|
|[<span data-ttu-id="7c6be-819">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-820">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-821">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-822">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-823">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-823">Returns:</span></span>

<span data-ttu-id="7c6be-824">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="7c6be-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="7c6be-825">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-825">Example</span></span>

<span data-ttu-id="7c6be-826">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="7c6be-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="7c6be-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="7c6be-828">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-829">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-830">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-830">Parameters</span></span>

|<span data-ttu-id="7c6be-831">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-831">Name</span></span>| <span data-ttu-id="7c6be-832">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-832">Type</span></span>| <span data-ttu-id="7c6be-833">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="7c6be-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="7c6be-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="7c6be-835">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="7c6be-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c6be-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-836">Requirements</span></span>

|<span data-ttu-id="7c6be-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-837">Requirement</span></span>| <span data-ttu-id="7c6be-838">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-840">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-840">1.0</span></span>|
|[<span data-ttu-id="7c6be-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-842">Restrito</span><span class="sxs-lookup"><span data-stu-id="7c6be-842">Restricted</span></span>|
|[<span data-ttu-id="7c6be-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-844">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-845">Returns:</span></span>

<span data-ttu-id="7c6be-846">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="7c6be-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="7c6be-847">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="7c6be-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="7c6be-848">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="7c6be-849">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="7c6be-850">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="7c6be-850">Value of `entityType`</span></span> | <span data-ttu-id="7c6be-851">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="7c6be-851">Type of objects in returned array</span></span> | <span data-ttu-id="7c6be-852">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="7c6be-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="7c6be-853">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-853">String</span></span> | <span data-ttu-id="7c6be-854">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="7c6be-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="7c6be-855">Contato</span><span class="sxs-lookup"><span data-stu-id="7c6be-855">Contact</span></span> | <span data-ttu-id="7c6be-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7c6be-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="7c6be-857">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-857">String</span></span> | <span data-ttu-id="7c6be-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7c6be-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="7c6be-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7c6be-859">MeetingSuggestion</span></span> | <span data-ttu-id="7c6be-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7c6be-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="7c6be-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7c6be-861">PhoneNumber</span></span> | <span data-ttu-id="7c6be-862">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="7c6be-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="7c6be-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7c6be-863">TaskSuggestion</span></span> | <span data-ttu-id="7c6be-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7c6be-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="7c6be-865">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-865">String</span></span> | <span data-ttu-id="7c6be-866">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="7c6be-866">**Restricted**</span></span> |

<span data-ttu-id="7c6be-867">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="7c6be-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="7c6be-868">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-868">Example</span></span>

<span data-ttu-id="7c6be-869">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="7c6be-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="7c6be-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="7c6be-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="7c6be-871">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-872">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7c6be-873">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-874">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-874">Parameters</span></span>

|<span data-ttu-id="7c6be-875">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-875">Name</span></span>| <span data-ttu-id="7c6be-876">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-876">Type</span></span>| <span data-ttu-id="7c6be-877">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7c6be-878">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-878">String</span></span>|<span data-ttu-id="7c6be-879">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="7c6be-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c6be-880">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-880">Requirements</span></span>

|<span data-ttu-id="7c6be-881">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-881">Requirement</span></span>| <span data-ttu-id="7c6be-882">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-883">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-884">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-884">1.0</span></span>|
|[<span data-ttu-id="7c6be-885">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-886">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-887">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-888">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-889">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-889">Returns:</span></span>

<span data-ttu-id="7c6be-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="7c6be-892">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="7c6be-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="7c6be-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7c6be-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="7c6be-894">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-895">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7c6be-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7c6be-899">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="7c6be-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7c6be-900">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="7c6be-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c6be-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-903">Requirements</span></span>

|<span data-ttu-id="7c6be-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-904">Requirement</span></span>| <span data-ttu-id="7c6be-905">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-906">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-907">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-907">1.0</span></span>|
|[<span data-ttu-id="7c6be-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-909">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-910">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-911">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-912">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-912">Returns:</span></span>

<span data-ttu-id="7c6be-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="7c6be-915">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="7c6be-916">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-916">Example</span></span>

<span data-ttu-id="7c6be-917">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="7c6be-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="7c6be-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="7c6be-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="7c6be-919">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7c6be-920">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="7c6be-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7c6be-921">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="7c6be-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-924">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-924">Parameters</span></span>

|<span data-ttu-id="7c6be-925">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-925">Name</span></span>| <span data-ttu-id="7c6be-926">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-926">Type</span></span>| <span data-ttu-id="7c6be-927">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7c6be-928">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-928">String</span></span>|<span data-ttu-id="7c6be-929">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="7c6be-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c6be-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-930">Requirements</span></span>

|<span data-ttu-id="7c6be-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-931">Requirement</span></span>| <span data-ttu-id="7c6be-932">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-934">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-934">1.0</span></span>|
|[<span data-ttu-id="7c6be-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-936">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-938">Read</span><span class="sxs-lookup"><span data-stu-id="7c6be-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-939">Returns:</span></span>

<span data-ttu-id="7c6be-940">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="7c6be-941">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="7c6be-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="7c6be-942">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="7c6be-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="7c6be-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="7c6be-944">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="7c6be-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna uma cadeia de caracteres vazia para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p165">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-947">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-947">Parameters</span></span>

|<span data-ttu-id="7c6be-948">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-948">Name</span></span>| <span data-ttu-id="7c6be-949">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-949">Type</span></span>| <span data-ttu-id="7c6be-950">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-950">Attributes</span></span>| <span data-ttu-id="7c6be-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-951">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="7c6be-952">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7c6be-952">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="7c6be-p166">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="7c6be-956">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-956">Object</span></span>| <span data-ttu-id="7c6be-957">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-957">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-958">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-958">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7c6be-959">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-959">Object</span></span>| <span data-ttu-id="7c6be-960">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-960">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-961">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-961">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7c6be-962">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-962">function</span></span>||<span data-ttu-id="7c6be-963">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-963">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7c6be-964">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-964">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="7c6be-965">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-965">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c6be-966">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-966">Requirements</span></span>

|<span data-ttu-id="7c6be-967">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-967">Requirement</span></span>| <span data-ttu-id="7c6be-968">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-968">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-969">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-969">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-970">1.2</span><span class="sxs-lookup"><span data-stu-id="7c6be-970">1.2</span></span>|
|[<span data-ttu-id="7c6be-971">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-971">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-972">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-972">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-973">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-973">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-974">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-974">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="7c6be-975">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7c6be-975">Returns:</span></span>

<span data-ttu-id="7c6be-976">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-976">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="7c6be-977">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="7c6be-977">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7c6be-978">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-978">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="7c6be-979">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7c6be-979">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="7c6be-980">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-980">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="7c6be-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-984">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-984">Parameters</span></span>

|<span data-ttu-id="7c6be-985">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-985">Name</span></span>| <span data-ttu-id="7c6be-986">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-986">Type</span></span>| <span data-ttu-id="7c6be-987">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-987">Attributes</span></span>| <span data-ttu-id="7c6be-988">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-988">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7c6be-989">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-989">function</span></span>||<span data-ttu-id="7c6be-990">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7c6be-991">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-991">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="7c6be-992">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="7c6be-992">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="7c6be-993">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-993">Object</span></span>| <span data-ttu-id="7c6be-994">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-994">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-995">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-995">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="7c6be-996">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-996">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c6be-997">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-997">Requirements</span></span>

|<span data-ttu-id="7c6be-998">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-998">Requirement</span></span>| <span data-ttu-id="7c6be-999">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-1000">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="7c6be-1001">1.0</span></span>|
|[<span data-ttu-id="7c6be-1002">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-1003">ReadItem</span></span>|
|[<span data-ttu-id="7c6be-1004">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7c6be-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-1005">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7c6be-1005">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-1006">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1006">Example</span></span>

<span data-ttu-id="7c6be-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="7c6be-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7c6be-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="7c6be-1011">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1011">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="7c6be-1012">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1012">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="7c6be-1013">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1013">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="7c6be-1014">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1014">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="7c6be-1015">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1015">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-1016">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-1016">Parameters</span></span>

|<span data-ttu-id="7c6be-1017">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-1017">Name</span></span>| <span data-ttu-id="7c6be-1018">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1018">Type</span></span>| <span data-ttu-id="7c6be-1019">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-1019">Attributes</span></span>| <span data-ttu-id="7c6be-1020">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-1020">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="7c6be-1021">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-1021">String</span></span>||<span data-ttu-id="7c6be-1022">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1022">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="7c6be-1023">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-1023">Object</span></span>| <span data-ttu-id="7c6be-1024">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1024">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1025">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1025">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7c6be-1026">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-1026">Object</span></span>| <span data-ttu-id="7c6be-1027">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1027">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1028">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1028">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7c6be-1029">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-1029">function</span></span>| <span data-ttu-id="7c6be-1030">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1030">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1031">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-1031">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7c6be-1032">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1032">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7c6be-1033">Erros</span><span class="sxs-lookup"><span data-stu-id="7c6be-1033">Errors</span></span>

| <span data-ttu-id="7c6be-1034">Código de erro</span><span class="sxs-lookup"><span data-stu-id="7c6be-1034">Error code</span></span> | <span data-ttu-id="7c6be-1035">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-1035">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="7c6be-1036">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1036">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-1037">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-1037">Requirements</span></span>

|<span data-ttu-id="7c6be-1038">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-1038">Requirement</span></span>| <span data-ttu-id="7c6be-1039">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-1039">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-1040">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-1040">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-1041">1.1</span><span class="sxs-lookup"><span data-stu-id="7c6be-1041">1.1</span></span>|
|[<span data-ttu-id="7c6be-1042">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1042">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-1043">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-1043">ReadWriteItem</span></span>|
|[<span data-ttu-id="7c6be-1044">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-1044">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-1045">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-1045">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-1046">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1046">Example</span></span>

<span data-ttu-id="7c6be-1047">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1047">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="7c6be-1048">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="7c6be-1048">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="7c6be-1049">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1049">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="7c6be-p173">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7c6be-1053">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7c6be-1053">Parameters</span></span>

|<span data-ttu-id="7c6be-1054">Nome</span><span class="sxs-lookup"><span data-stu-id="7c6be-1054">Name</span></span>| <span data-ttu-id="7c6be-1055">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1055">Type</span></span>| <span data-ttu-id="7c6be-1056">Atributos</span><span class="sxs-lookup"><span data-stu-id="7c6be-1056">Attributes</span></span>| <span data-ttu-id="7c6be-1057">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c6be-1057">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7c6be-1058">String</span><span class="sxs-lookup"><span data-stu-id="7c6be-1058">String</span></span>||<span data-ttu-id="7c6be-p174">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="7c6be-1062">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-1062">Object</span></span>| <span data-ttu-id="7c6be-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1064">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7c6be-1065">Objeto</span><span class="sxs-lookup"><span data-stu-id="7c6be-1065">Object</span></span>| <span data-ttu-id="7c6be-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1067">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="7c6be-1068">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7c6be-1068">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="7c6be-1069">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7c6be-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="7c6be-1070">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1070">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="7c6be-1071">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1071">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="7c6be-1072">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1072">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="7c6be-1073">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1073">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="7c6be-1074">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="7c6be-1074">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="7c6be-1075">function</span><span class="sxs-lookup"><span data-stu-id="7c6be-1075">function</span></span>||<span data-ttu-id="7c6be-1076">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7c6be-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c6be-1077">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c6be-1077">Requirements</span></span>

|<span data-ttu-id="7c6be-1078">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c6be-1078">Requirement</span></span>| <span data-ttu-id="7c6be-1079">Valor</span><span class="sxs-lookup"><span data-stu-id="7c6be-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c6be-1080">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c6be-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c6be-1081">1.2</span><span class="sxs-lookup"><span data-stu-id="7c6be-1081">1.2</span></span>|
|[<span data-ttu-id="7c6be-1082">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c6be-1083">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7c6be-1083">ReadWriteItem</span></span>|
|[<span data-ttu-id="7c6be-1084">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c6be-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7c6be-1085">Escrever</span><span class="sxs-lookup"><span data-stu-id="7c6be-1085">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7c6be-1086">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c6be-1086">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
