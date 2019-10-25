---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,1
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 3d0b9783ea7fd235f4f989182ced04e0bce735ff
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682652"
---
# <a name="item"></a><span data-ttu-id="31fe9-102">item</span><span class="sxs-lookup"><span data-stu-id="31fe9-102">item</span></span>

### <span data-ttu-id="31fe9-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="31fe9-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="31fe9-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="31fe9-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-107">Requirements</span></span>

|<span data-ttu-id="31fe9-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-108">Requirement</span></span>| <span data-ttu-id="31fe9-109">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-111">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-111">1.0</span></span>|
|[<span data-ttu-id="31fe9-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="31fe9-113">Restricted</span></span>|
|[<span data-ttu-id="31fe9-114">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-115">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="31fe9-116">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="31fe9-116">Members and methods</span></span>

| <span data-ttu-id="31fe9-117">Membro	</span><span class="sxs-lookup"><span data-stu-id="31fe9-117">Member</span></span> | <span data-ttu-id="31fe9-118">Tipo	</span><span class="sxs-lookup"><span data-stu-id="31fe9-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="31fe9-119">attachments</span><span class="sxs-lookup"><span data-stu-id="31fe9-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="31fe9-120">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-120">Member</span></span> |
| [<span data-ttu-id="31fe9-121">bcc</span><span class="sxs-lookup"><span data-stu-id="31fe9-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="31fe9-122">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-122">Member</span></span> |
| [<span data-ttu-id="31fe9-123">body</span><span class="sxs-lookup"><span data-stu-id="31fe9-123">body</span></span>](#body-body) | <span data-ttu-id="31fe9-124">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-124">Member</span></span> |
| [<span data-ttu-id="31fe9-125">cc</span><span class="sxs-lookup"><span data-stu-id="31fe9-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="31fe9-126">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-126">Member</span></span> |
| [<span data-ttu-id="31fe9-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="31fe9-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="31fe9-128">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-128">Member</span></span> |
| [<span data-ttu-id="31fe9-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="31fe9-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="31fe9-130">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-130">Member</span></span> |
| [<span data-ttu-id="31fe9-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="31fe9-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="31fe9-132">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-132">Member</span></span> |
| [<span data-ttu-id="31fe9-133">end</span><span class="sxs-lookup"><span data-stu-id="31fe9-133">end</span></span>](#end-datetime) | <span data-ttu-id="31fe9-134">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-134">Member</span></span> |
| [<span data-ttu-id="31fe9-135">from</span><span class="sxs-lookup"><span data-stu-id="31fe9-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="31fe9-136">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-136">Member</span></span> |
| [<span data-ttu-id="31fe9-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="31fe9-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="31fe9-138">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-138">Member</span></span> |
| [<span data-ttu-id="31fe9-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="31fe9-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="31fe9-140">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-140">Member</span></span> |
| [<span data-ttu-id="31fe9-141">itemId</span><span class="sxs-lookup"><span data-stu-id="31fe9-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="31fe9-142">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-142">Member</span></span> |
| [<span data-ttu-id="31fe9-143">itemType</span><span class="sxs-lookup"><span data-stu-id="31fe9-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="31fe9-144">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-144">Member</span></span> |
| [<span data-ttu-id="31fe9-145">location</span><span class="sxs-lookup"><span data-stu-id="31fe9-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="31fe9-146">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-146">Member</span></span> |
| [<span data-ttu-id="31fe9-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="31fe9-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="31fe9-148">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-148">Member</span></span> |
| [<span data-ttu-id="31fe9-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="31fe9-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="31fe9-150">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-150">Member</span></span> |
| [<span data-ttu-id="31fe9-151">organizer</span><span class="sxs-lookup"><span data-stu-id="31fe9-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="31fe9-152">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-152">Member</span></span> |
| [<span data-ttu-id="31fe9-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="31fe9-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="31fe9-154">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-154">Member</span></span> |
| [<span data-ttu-id="31fe9-155">sender</span><span class="sxs-lookup"><span data-stu-id="31fe9-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="31fe9-156">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-156">Member</span></span> |
| [<span data-ttu-id="31fe9-157">start</span><span class="sxs-lookup"><span data-stu-id="31fe9-157">start</span></span>](#start-datetime) | <span data-ttu-id="31fe9-158">Member</span><span class="sxs-lookup"><span data-stu-id="31fe9-158">Member</span></span> |
| [<span data-ttu-id="31fe9-159">subject</span><span class="sxs-lookup"><span data-stu-id="31fe9-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="31fe9-160">Membro</span><span class="sxs-lookup"><span data-stu-id="31fe9-160">Member</span></span> |
| [<span data-ttu-id="31fe9-161">to</span><span class="sxs-lookup"><span data-stu-id="31fe9-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="31fe9-162">Membro</span><span class="sxs-lookup"><span data-stu-id="31fe9-162">Member</span></span> |
| [<span data-ttu-id="31fe9-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="31fe9-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="31fe9-164">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-164">Method</span></span> |
| [<span data-ttu-id="31fe9-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="31fe9-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="31fe9-166">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-166">Method</span></span> |
| [<span data-ttu-id="31fe9-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="31fe9-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="31fe9-168">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-168">Method</span></span> |
| [<span data-ttu-id="31fe9-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="31fe9-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="31fe9-170">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-170">Method</span></span> |
| [<span data-ttu-id="31fe9-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="31fe9-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="31fe9-172">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-172">Method</span></span> |
| [<span data-ttu-id="31fe9-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="31fe9-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="31fe9-174">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-174">Method</span></span> |
| [<span data-ttu-id="31fe9-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="31fe9-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="31fe9-176">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-176">Method</span></span> |
| [<span data-ttu-id="31fe9-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="31fe9-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="31fe9-178">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-178">Method</span></span> |
| [<span data-ttu-id="31fe9-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="31fe9-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="31fe9-180">Method</span><span class="sxs-lookup"><span data-stu-id="31fe9-180">Method</span></span> |
| [<span data-ttu-id="31fe9-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="31fe9-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="31fe9-182">Método</span><span class="sxs-lookup"><span data-stu-id="31fe9-182">Method</span></span> |
| [<span data-ttu-id="31fe9-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="31fe9-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="31fe9-184">Método</span><span class="sxs-lookup"><span data-stu-id="31fe9-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="31fe9-185">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-185">Example</span></span>

<span data-ttu-id="31fe9-186">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="31fe9-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="31fe9-187">Members</span><span class="sxs-lookup"><span data-stu-id="31fe9-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="31fe9-188">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="31fe9-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="31fe9-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-191">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="31fe9-192">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="31fe9-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-193">Type</span></span>

*   <span data-ttu-id="31fe9-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="31fe9-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-195">Requirements</span></span>

|<span data-ttu-id="31fe9-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-196">Requirement</span></span>| <span data-ttu-id="31fe9-197">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-199">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-199">1.0</span></span>|
|[<span data-ttu-id="31fe9-200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-201">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-203">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-204">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-204">Example</span></span>

<span data-ttu-id="31fe9-205">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="31fe9-206">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-207">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="31fe9-208">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="31fe9-208">Compose mode only.</span></span>

<span data-ttu-id="31fe9-209">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-209">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-210">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-210">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="31fe9-211">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-211">Get 500 members maximum.</span></span>
- <span data-ttu-id="31fe9-212">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="31fe9-212">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-213">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-213">Type</span></span>

*   [<span data-ttu-id="31fe9-214">Destinatários</span><span class="sxs-lookup"><span data-stu-id="31fe9-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-215">Requirements</span></span>

|<span data-ttu-id="31fe9-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-216">Requirement</span></span>| <span data-ttu-id="31fe9-217">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-219">1.1</span><span class="sxs-lookup"><span data-stu-id="31fe9-219">1.1</span></span>|
|[<span data-ttu-id="31fe9-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-221">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-223">Escrever</span><span class="sxs-lookup"><span data-stu-id="31fe9-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="31fe9-225">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-226">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="31fe9-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-227">Type</span></span>

*   [<span data-ttu-id="31fe9-228">Body</span><span class="sxs-lookup"><span data-stu-id="31fe9-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-229">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-229">Requirements</span></span>

|<span data-ttu-id="31fe9-230">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-230">Requirement</span></span>| <span data-ttu-id="31fe9-231">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-232">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-233">1.1</span><span class="sxs-lookup"><span data-stu-id="31fe9-233">1.1</span></span>|
|[<span data-ttu-id="31fe9-234">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-235">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-236">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-238">Example</span></span>

<span data-ttu-id="31fe9-239">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="31fe9-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="31fe9-240">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="31fe9-241">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-242">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="31fe9-243">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-244">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-244">Read mode</span></span>

<span data-ttu-id="31fe9-245">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-245">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="31fe9-246">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-246">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-247">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-247">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-248">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-248">Compose mode</span></span>

<span data-ttu-id="31fe9-249">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="31fe9-250">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-251">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="31fe9-252">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="31fe9-253">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="31fe9-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-254">Type</span></span>

*   <span data-ttu-id="31fe9-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-256">Requirements</span></span>

|<span data-ttu-id="31fe9-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-257">Requirement</span></span>| <span data-ttu-id="31fe9-258">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-260">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-260">1.0</span></span>|
|[<span data-ttu-id="31fe9-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-262">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-263">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="31fe9-265">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="31fe9-266">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="31fe9-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="31fe9-p110">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="31fe9-p111">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-271">Type</span></span>

*   <span data-ttu-id="31fe9-272">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-273">Requirements</span></span>

|<span data-ttu-id="31fe9-274">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-274">Requirement</span></span>| <span data-ttu-id="31fe9-275">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-277">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-277">1.0</span></span>|
|[<span data-ttu-id="31fe9-278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-279">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-281">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="31fe9-283">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="31fe9-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="31fe9-p112">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-286">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-286">Type</span></span>

*   <span data-ttu-id="31fe9-287">Data</span><span class="sxs-lookup"><span data-stu-id="31fe9-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-288">Requirements</span></span>

|<span data-ttu-id="31fe9-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-289">Requirement</span></span>| <span data-ttu-id="31fe9-290">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-292">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-292">1.0</span></span>|
|[<span data-ttu-id="31fe9-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-294">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-295">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-296">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="31fe9-298">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="31fe9-298">dateTimeModified: Date</span></span>

<span data-ttu-id="31fe9-p113">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-301">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-302">Type</span></span>

*   <span data-ttu-id="31fe9-303">Data</span><span class="sxs-lookup"><span data-stu-id="31fe9-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-304">Requirements</span></span>

|<span data-ttu-id="31fe9-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-305">Requirement</span></span>| <span data-ttu-id="31fe9-306">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-308">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-308">1.0</span></span>|
|[<span data-ttu-id="31fe9-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-310">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-311">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-312">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="31fe9-314">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="31fe9-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="31fe9-p114">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-318">Read mode</span></span>

<span data-ttu-id="31fe9-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-320">Compose mode</span></span>

<span data-ttu-id="31fe9-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="31fe9-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="31fe9-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="31fe9-323">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="31fe9-324">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-324">Type</span></span>

*   <span data-ttu-id="31fe9-325">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-326">Requirements</span></span>

|<span data-ttu-id="31fe9-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-327">Requirement</span></span>| <span data-ttu-id="31fe9-328">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-330">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-330">1.0</span></span>|
|[<span data-ttu-id="31fe9-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-332">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-333">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="31fe9-335">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-p115">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="31fe9-p116">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-340">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-341">Type</span></span>

*   [<span data-ttu-id="31fe9-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="31fe9-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-343">Requirements</span></span>

|<span data-ttu-id="31fe9-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-344">Requirement</span></span>| <span data-ttu-id="31fe9-345">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-347">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-347">1.0</span></span>|
|[<span data-ttu-id="31fe9-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-349">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-351">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="31fe9-353">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-353">internetMessageId: String</span></span>

<span data-ttu-id="31fe9-p117">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-356">Type</span></span>

*   <span data-ttu-id="31fe9-357">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-358">Requirements</span></span>

|<span data-ttu-id="31fe9-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-359">Requirement</span></span>| <span data-ttu-id="31fe9-360">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-362">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-362">1.0</span></span>|
|[<span data-ttu-id="31fe9-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-364">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-366">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="31fe9-368">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="31fe9-368">itemClass: String</span></span>

<span data-ttu-id="31fe9-p118">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="31fe9-p119">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="31fe9-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-373">Type</span></span> | <span data-ttu-id="31fe9-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-374">Description</span></span> | <span data-ttu-id="31fe9-375">classe de item</span><span class="sxs-lookup"><span data-stu-id="31fe9-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="31fe9-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="31fe9-376">Appointment items</span></span> | <span data-ttu-id="31fe9-377">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="31fe9-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="31fe9-378">Message items</span></span> | <span data-ttu-id="31fe9-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="31fe9-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="31fe9-380">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-381">Type</span></span>

*   <span data-ttu-id="31fe9-382">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-383">Requirements</span></span>

|<span data-ttu-id="31fe9-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-384">Requirement</span></span>| <span data-ttu-id="31fe9-385">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-387">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-387">1.0</span></span>|
|[<span data-ttu-id="31fe9-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-389">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-391">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="31fe9-393">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-393">(nullable) itemId: String</span></span>

<span data-ttu-id="31fe9-p120">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p120">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="31fe9-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="31fe9-397">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="31fe9-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="31fe9-398">Antes de fazer chamadas da API REST usando esse valor, ele deve ser `Office.context.mailbox.convertToRestId`convertido usando o, que está disponível a partir do conjunto de requisitos 1,3.</span><span class="sxs-lookup"><span data-stu-id="31fe9-398">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="31fe9-399">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="31fe9-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-400">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-400">Type</span></span>

*   <span data-ttu-id="31fe9-401">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-402">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-402">Requirements</span></span>

|<span data-ttu-id="31fe9-403">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-403">Requirement</span></span>| <span data-ttu-id="31fe9-404">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-405">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-406">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-406">1.0</span></span>|
|[<span data-ttu-id="31fe9-407">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-407">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-408">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-409">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-409">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-410">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-411">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-411">Example</span></span>

<span data-ttu-id="31fe9-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="31fe9-414">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-415">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="31fe9-415">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="31fe9-416">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-416">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-417">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-417">Type</span></span>

*   [<span data-ttu-id="31fe9-418">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="31fe9-418">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-419">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-419">Requirements</span></span>

|<span data-ttu-id="31fe9-420">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-420">Requirement</span></span>| <span data-ttu-id="31fe9-421">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-421">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-422">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-423">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-423">1.0</span></span>|
|[<span data-ttu-id="31fe9-424">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-425">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-425">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-426">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-427">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-427">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-428">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-428">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="31fe9-429">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-430">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-430">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-431">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-431">Read mode</span></span>

<span data-ttu-id="31fe9-432">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-432">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-433">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-433">Compose mode</span></span>

<span data-ttu-id="31fe9-434">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-434">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-435">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-435">Type</span></span>

*   <span data-ttu-id="31fe9-436">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-437">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-437">Requirements</span></span>

|<span data-ttu-id="31fe9-438">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-438">Requirement</span></span>| <span data-ttu-id="31fe9-439">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-440">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-441">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-441">1.0</span></span>|
|[<span data-ttu-id="31fe9-442">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-443">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-444">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-445">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-445">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="31fe9-446">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-446">normalizedSubject: String</span></span>

<span data-ttu-id="31fe9-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="31fe9-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="31fe9-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-451">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-451">Type</span></span>

*   <span data-ttu-id="31fe9-452">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-453">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-453">Requirements</span></span>

|<span data-ttu-id="31fe9-454">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-454">Requirement</span></span>| <span data-ttu-id="31fe9-455">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-456">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-457">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-457">1.0</span></span>|
|[<span data-ttu-id="31fe9-458">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-459">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-460">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-461">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-462">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="31fe9-463">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-464">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="31fe9-464">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="31fe9-465">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-465">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-466">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-466">Read mode</span></span>

<span data-ttu-id="31fe9-467">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="31fe9-467">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="31fe9-468">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-468">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-469">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-469">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-470">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-470">Compose mode</span></span>

<span data-ttu-id="31fe9-471">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="31fe9-471">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="31fe9-472">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-473">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-473">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="31fe9-474">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-474">Get 500 members maximum.</span></span>
- <span data-ttu-id="31fe9-475">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="31fe9-475">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-476">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-476">Type</span></span>

*   <span data-ttu-id="31fe9-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-478">Requirements</span></span>

|<span data-ttu-id="31fe9-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-479">Requirement</span></span>| <span data-ttu-id="31fe9-480">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-482">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-482">1.0</span></span>|
|[<span data-ttu-id="31fe9-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-484">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-485">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-486">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-486">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="31fe9-487">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-p128">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-490">Type</span></span>

*   [<span data-ttu-id="31fe9-491">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="31fe9-491">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-492">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-492">Requirements</span></span>

|<span data-ttu-id="31fe9-493">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-493">Requirement</span></span>| <span data-ttu-id="31fe9-494">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-495">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-496">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-496">1.0</span></span>|
|[<span data-ttu-id="31fe9-497">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-498">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-499">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-500">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-500">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-501">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-501">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="31fe9-502">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-503">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="31fe9-503">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="31fe9-504">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-504">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-505">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-505">Read mode</span></span>

<span data-ttu-id="31fe9-506">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="31fe9-506">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="31fe9-507">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-507">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-508">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-508">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-509">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-509">Compose mode</span></span>

<span data-ttu-id="31fe9-510">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="31fe9-510">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="31fe9-511">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-512">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-512">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="31fe9-513">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-513">Get 500 members maximum.</span></span>
- <span data-ttu-id="31fe9-514">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="31fe9-514">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-515">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-515">Type</span></span>

*   <span data-ttu-id="31fe9-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-517">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-517">Requirements</span></span>

|<span data-ttu-id="31fe9-518">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-518">Requirement</span></span>| <span data-ttu-id="31fe9-519">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-520">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-521">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-521">1.0</span></span>|
|[<span data-ttu-id="31fe9-522">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-523">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-524">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-525">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-525">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="31fe9-526">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-p132">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="31fe9-p133">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-531">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-531">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="31fe9-532">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-532">Type</span></span>

*   [<span data-ttu-id="31fe9-533">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="31fe9-533">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="31fe9-534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-534">Requirements</span></span>

|<span data-ttu-id="31fe9-535">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-535">Requirement</span></span>| <span data-ttu-id="31fe9-536">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-537">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-538">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-538">1.0</span></span>|
|[<span data-ttu-id="31fe9-539">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-540">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-541">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-542">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-542">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-543">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-543">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="31fe9-544">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-545">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="31fe9-545">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="31fe9-p134">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-548">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-548">Read mode</span></span>

<span data-ttu-id="31fe9-549">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-549">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-550">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-550">Compose mode</span></span>

<span data-ttu-id="31fe9-551">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-551">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="31fe9-552">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="31fe9-552">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="31fe9-553">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-553">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="31fe9-554">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-554">Type</span></span>

*   <span data-ttu-id="31fe9-555">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-556">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-556">Requirements</span></span>

|<span data-ttu-id="31fe9-557">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-557">Requirement</span></span>| <span data-ttu-id="31fe9-558">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-559">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-560">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-560">1.0</span></span>|
|[<span data-ttu-id="31fe9-561">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-562">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-563">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-564">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-564">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="31fe9-565">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-566">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="31fe9-566">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="31fe9-567">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="31fe9-567">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-568">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-568">Read mode</span></span>

<span data-ttu-id="31fe9-p135">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-571">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-571">Compose mode</span></span>

<span data-ttu-id="31fe9-572">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="31fe9-572">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-573">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-573">Type</span></span>

*   <span data-ttu-id="31fe9-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-575">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-575">Requirements</span></span>

|<span data-ttu-id="31fe9-576">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-576">Requirement</span></span>| <span data-ttu-id="31fe9-577">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-577">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-578">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-578">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-579">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-579">1.0</span></span>|
|[<span data-ttu-id="31fe9-580">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-580">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-581">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-581">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-582">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-582">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-583">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-583">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="31fe9-584">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="31fe9-585">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-585">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="31fe9-586">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-586">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="31fe9-587">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="31fe9-587">Read mode</span></span>

<span data-ttu-id="31fe9-588">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-588">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="31fe9-589">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-589">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-590">No entanto, no Windows e no Mac, é possível obter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-590">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="31fe9-591">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="31fe9-591">Compose mode</span></span>

<span data-ttu-id="31fe9-592">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31fe9-592">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="31fe9-593">Por padrão, a coleção é limitada a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="31fe9-594">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="31fe9-594">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="31fe9-595">Obter máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="31fe9-595">Get 500 members maximum.</span></span>
- <span data-ttu-id="31fe9-596">Defina um máximo de 100 membros por chamada, até 500, no total.</span><span class="sxs-lookup"><span data-stu-id="31fe9-596">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="31fe9-597">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-597">Type</span></span>

*   <span data-ttu-id="31fe9-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-599">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-599">Requirements</span></span>

|<span data-ttu-id="31fe9-600">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-600">Requirement</span></span>| <span data-ttu-id="31fe9-601">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-602">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-603">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-603">1.0</span></span>|
|[<span data-ttu-id="31fe9-604">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-605">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-606">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-607">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-607">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="31fe9-608">Métodos</span><span class="sxs-lookup"><span data-stu-id="31fe9-608">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="31fe9-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="31fe9-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="31fe9-610">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="31fe9-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="31fe9-611">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="31fe9-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="31fe9-612">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="31fe9-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-613">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-613">Parameters</span></span>

|<span data-ttu-id="31fe9-614">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-614">Name</span></span>| <span data-ttu-id="31fe9-615">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-615">Type</span></span>| <span data-ttu-id="31fe9-616">Atributos</span><span class="sxs-lookup"><span data-stu-id="31fe9-616">Attributes</span></span>| <span data-ttu-id="31fe9-617">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="31fe9-618">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-618">String</span></span>||<span data-ttu-id="31fe9-p139">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="31fe9-621">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-621">String</span></span>||<span data-ttu-id="31fe9-p140">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="31fe9-624">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-624">Object</span></span>| <span data-ttu-id="31fe9-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-625">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-626">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-626">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="31fe9-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-627">Object</span></span>| <span data-ttu-id="31fe9-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-628">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-629">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-629">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="31fe9-630">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-630">function</span></span>| <span data-ttu-id="31fe9-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-631">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-632">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="31fe9-633">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-633">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="31fe9-634">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="31fe9-634">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="31fe9-635">Erros</span><span class="sxs-lookup"><span data-stu-id="31fe9-635">Errors</span></span>

| <span data-ttu-id="31fe9-636">Código de erro</span><span class="sxs-lookup"><span data-stu-id="31fe9-636">Error code</span></span> | <span data-ttu-id="31fe9-637">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-637">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="31fe9-638">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="31fe9-638">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="31fe9-639">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="31fe9-639">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="31fe9-640">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="31fe9-640">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31fe9-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-641">Requirements</span></span>

|<span data-ttu-id="31fe9-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-642">Requirement</span></span>| <span data-ttu-id="31fe9-643">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-644">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-645">1.1</span><span class="sxs-lookup"><span data-stu-id="31fe9-645">1.1</span></span>|
|[<span data-ttu-id="31fe9-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-647">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-647">ReadWriteItem</span></span>|
|[<span data-ttu-id="31fe9-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-649">Escrever</span><span class="sxs-lookup"><span data-stu-id="31fe9-649">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-650">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-650">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="31fe9-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="31fe9-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="31fe9-652">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-652">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="31fe9-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="31fe9-656">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="31fe9-656">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="31fe9-657">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-657">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-658">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-658">Parameters</span></span>

|<span data-ttu-id="31fe9-659">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-659">Name</span></span>| <span data-ttu-id="31fe9-660">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-660">Type</span></span>| <span data-ttu-id="31fe9-661">Atributos</span><span class="sxs-lookup"><span data-stu-id="31fe9-661">Attributes</span></span>| <span data-ttu-id="31fe9-662">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-662">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="31fe9-663">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-663">String</span></span>||<span data-ttu-id="31fe9-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="31fe9-666">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-666">String</span></span>||<span data-ttu-id="31fe9-667">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-667">The subject of the item to be attached.</span></span> <span data-ttu-id="31fe9-668">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31fe9-668">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="31fe9-669">Object</span><span class="sxs-lookup"><span data-stu-id="31fe9-669">Object</span></span>| <span data-ttu-id="31fe9-670">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-670">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-671">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-671">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="31fe9-672">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-672">Object</span></span>| <span data-ttu-id="31fe9-673">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-673">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-674">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-674">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="31fe9-675">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-675">function</span></span>| <span data-ttu-id="31fe9-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-676">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-677">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="31fe9-678">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-678">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="31fe9-679">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="31fe9-679">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="31fe9-680">Erros</span><span class="sxs-lookup"><span data-stu-id="31fe9-680">Errors</span></span>

| <span data-ttu-id="31fe9-681">Código de erro</span><span class="sxs-lookup"><span data-stu-id="31fe9-681">Error code</span></span> | <span data-ttu-id="31fe9-682">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-682">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="31fe9-683">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="31fe9-683">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31fe9-684">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-684">Requirements</span></span>

|<span data-ttu-id="31fe9-685">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-685">Requirement</span></span>| <span data-ttu-id="31fe9-686">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-686">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-687">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-687">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-688">1.1</span><span class="sxs-lookup"><span data-stu-id="31fe9-688">1.1</span></span>|
|[<span data-ttu-id="31fe9-689">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-689">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-690">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-690">ReadWriteItem</span></span>|
|[<span data-ttu-id="31fe9-691">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-691">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-692">Escrever</span><span class="sxs-lookup"><span data-stu-id="31fe9-692">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-693">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-693">Example</span></span>

<span data-ttu-id="31fe9-694">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-694">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="31fe9-695">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="31fe9-695">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="31fe9-696">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-696">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-697">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-697">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="31fe9-698">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="31fe9-698">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="31fe9-699">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="31fe9-699">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-700">A capacidade de incluir anexos na chamada para `displayReplyAllForm` não é suportada no conjunto de requisitos 1,1.</span><span class="sxs-lookup"><span data-stu-id="31fe9-700">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="31fe9-701">O suporte a anexos foi adicionado a `displayReplyAllForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="31fe9-701">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-702">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-702">Parameters</span></span>

|<span data-ttu-id="31fe9-703">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-703">Name</span></span>| <span data-ttu-id="31fe9-704">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-704">Type</span></span>| <span data-ttu-id="31fe9-705">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-705">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="31fe9-706">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="31fe9-706">String &#124; Object</span></span>| |<span data-ttu-id="31fe9-p145">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="31fe9-709">**OU**</span><span class="sxs-lookup"><span data-stu-id="31fe9-709">**OR**</span></span><br/><span data-ttu-id="31fe9-p146">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="31fe9-712">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-712">String</span></span> | <span data-ttu-id="31fe9-713">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-713">&lt;optional&gt;</span></span> | <span data-ttu-id="31fe9-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="31fe9-716">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-716">function</span></span> | <span data-ttu-id="31fe9-717">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-717">&lt;optional&gt;</span></span> | <span data-ttu-id="31fe9-718">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-718">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31fe9-719">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-719">Requirements</span></span>

|<span data-ttu-id="31fe9-720">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-720">Requirement</span></span>| <span data-ttu-id="31fe9-721">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-721">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-722">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-722">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-723">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-723">1.0</span></span>|
|[<span data-ttu-id="31fe9-724">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-724">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-725">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-725">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-726">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-726">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-727">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-727">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="31fe9-728">Exemplos</span><span class="sxs-lookup"><span data-stu-id="31fe9-728">Examples</span></span>

<span data-ttu-id="31fe9-729">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-729">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="31fe9-730">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="31fe9-730">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="31fe9-731">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="31fe9-731">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="31fe9-732">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-732">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="31fe9-733">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="31fe9-733">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="31fe9-734">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-734">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-735">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-735">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="31fe9-736">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="31fe9-736">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="31fe9-737">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="31fe9-737">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-738">A capacidade de incluir anexos na chamada para `displayReplyForm` não é suportada no conjunto de requisitos 1,1.</span><span class="sxs-lookup"><span data-stu-id="31fe9-738">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="31fe9-739">O suporte a anexos foi adicionado a `displayReplyForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="31fe9-739">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-740">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-740">Parameters</span></span>

|<span data-ttu-id="31fe9-741">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-741">Name</span></span>| <span data-ttu-id="31fe9-742">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-742">Type</span></span>| <span data-ttu-id="31fe9-743">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-743">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="31fe9-744">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="31fe9-744">String &#124; Object</span></span>| | <span data-ttu-id="31fe9-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="31fe9-747">**OU**</span><span class="sxs-lookup"><span data-stu-id="31fe9-747">**OR**</span></span><br/><span data-ttu-id="31fe9-p150">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p150">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="31fe9-750">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-750">String</span></span> | <span data-ttu-id="31fe9-751">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-751">&lt;optional&gt;</span></span> | <span data-ttu-id="31fe9-p151">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="31fe9-754">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-754">function</span></span> | <span data-ttu-id="31fe9-755">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-755">&lt;optional&gt;</span></span> | <span data-ttu-id="31fe9-756">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31fe9-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-757">Requirements</span></span>

|<span data-ttu-id="31fe9-758">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-758">Requirement</span></span>| <span data-ttu-id="31fe9-759">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-760">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-761">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-761">1.0</span></span>|
|[<span data-ttu-id="31fe9-762">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-763">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-764">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-765">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="31fe9-766">Exemplos</span><span class="sxs-lookup"><span data-stu-id="31fe9-766">Examples</span></span>

<span data-ttu-id="31fe9-767">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-767">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="31fe9-768">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="31fe9-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="31fe9-769">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="31fe9-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="31fe9-770">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-770">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="31fe9-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="31fe9-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="31fe9-772">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-772">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-773">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-773">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-774">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-774">Requirements</span></span>

|<span data-ttu-id="31fe9-775">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-775">Requirement</span></span>| <span data-ttu-id="31fe9-776">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-776">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-777">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-777">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-778">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-778">1.0</span></span>|
|[<span data-ttu-id="31fe9-779">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-779">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-780">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-780">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-781">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-781">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-782">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-782">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31fe9-783">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31fe9-783">Returns:</span></span>

<span data-ttu-id="31fe9-784">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="31fe9-784">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="31fe9-785">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-785">Example</span></span>

<span data-ttu-id="31fe9-786">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-786">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="31fe9-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="31fe9-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="31fe9-788">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-788">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-789">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-789">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-790">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-790">Parameters</span></span>

|<span data-ttu-id="31fe9-791">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-791">Name</span></span>| <span data-ttu-id="31fe9-792">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-792">Type</span></span>| <span data-ttu-id="31fe9-793">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-793">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="31fe9-794">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="31fe9-794">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="31fe9-795">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="31fe9-795">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31fe9-796">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-796">Requirements</span></span>

|<span data-ttu-id="31fe9-797">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-797">Requirement</span></span>| <span data-ttu-id="31fe9-798">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-799">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-800">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-800">1.0</span></span>|
|[<span data-ttu-id="31fe9-801">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-802">Restrito</span><span class="sxs-lookup"><span data-stu-id="31fe9-802">Restricted</span></span>|
|[<span data-ttu-id="31fe9-803">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-804">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31fe9-805">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31fe9-805">Returns:</span></span>

<span data-ttu-id="31fe9-806">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="31fe9-806">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="31fe9-807">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="31fe9-807">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="31fe9-808">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-808">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="31fe9-809">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-809">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="31fe9-810">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="31fe9-810">Value of `entityType`</span></span> | <span data-ttu-id="31fe9-811">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="31fe9-811">Type of objects in returned array</span></span> | <span data-ttu-id="31fe9-812">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="31fe9-812">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="31fe9-813">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-813">String</span></span> | <span data-ttu-id="31fe9-814">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="31fe9-814">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="31fe9-815">Contato</span><span class="sxs-lookup"><span data-stu-id="31fe9-815">Contact</span></span> | <span data-ttu-id="31fe9-816">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="31fe9-816">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="31fe9-817">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-817">String</span></span> | <span data-ttu-id="31fe9-818">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="31fe9-818">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="31fe9-819">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="31fe9-819">MeetingSuggestion</span></span> | <span data-ttu-id="31fe9-820">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="31fe9-820">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="31fe9-821">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="31fe9-821">PhoneNumber</span></span> | <span data-ttu-id="31fe9-822">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="31fe9-822">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="31fe9-823">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="31fe9-823">TaskSuggestion</span></span> | <span data-ttu-id="31fe9-824">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="31fe9-824">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="31fe9-825">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-825">String</span></span> | <span data-ttu-id="31fe9-826">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="31fe9-826">**Restricted**</span></span> |

<span data-ttu-id="31fe9-827">Tipo:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="31fe9-827">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="31fe9-828">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-828">Example</span></span>

<span data-ttu-id="31fe9-829">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="31fe9-829">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="31fe9-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="31fe9-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="31fe9-831">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="31fe9-831">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-832">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-832">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="31fe9-833">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-833">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-834">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-834">Parameters</span></span>

|<span data-ttu-id="31fe9-835">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-835">Name</span></span>| <span data-ttu-id="31fe9-836">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-836">Type</span></span>| <span data-ttu-id="31fe9-837">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-837">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="31fe9-838">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-838">String</span></span>|<span data-ttu-id="31fe9-839">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="31fe9-839">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31fe9-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-840">Requirements</span></span>

|<span data-ttu-id="31fe9-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-841">Requirement</span></span>| <span data-ttu-id="31fe9-842">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-844">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-844">1.0</span></span>|
|[<span data-ttu-id="31fe9-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-846">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-848">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31fe9-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31fe9-849">Returns:</span></span>

<span data-ttu-id="31fe9-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="31fe9-852">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="31fe9-852">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="31fe9-853">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="31fe9-853">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="31fe9-854">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="31fe9-854">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-855">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="31fe9-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="31fe9-859">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="31fe9-859">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="31fe9-860">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-860">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="31fe9-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31fe9-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-863">Requirements</span></span>

|<span data-ttu-id="31fe9-864">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-864">Requirement</span></span>| <span data-ttu-id="31fe9-865">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-866">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-867">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-867">1.0</span></span>|
|[<span data-ttu-id="31fe9-868">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-869">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-870">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-871">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31fe9-872">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31fe9-872">Returns:</span></span>

<span data-ttu-id="31fe9-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="31fe9-875">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-875">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="31fe9-876">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-876">Example</span></span>

<span data-ttu-id="31fe9-877">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="31fe9-877">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="31fe9-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="31fe9-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="31fe9-879">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="31fe9-879">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="31fe9-880">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="31fe9-880">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="31fe9-881">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-881">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="31fe9-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-884">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-884">Parameters</span></span>

|<span data-ttu-id="31fe9-885">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-885">Name</span></span>| <span data-ttu-id="31fe9-886">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-886">Type</span></span>| <span data-ttu-id="31fe9-887">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="31fe9-888">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="31fe9-888">String</span></span>|<span data-ttu-id="31fe9-889">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="31fe9-889">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31fe9-890">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-890">Requirements</span></span>

|<span data-ttu-id="31fe9-891">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-891">Requirement</span></span>| <span data-ttu-id="31fe9-892">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-893">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-894">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-894">1.0</span></span>|
|[<span data-ttu-id="31fe9-895">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-896">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-897">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-898">Read</span><span class="sxs-lookup"><span data-stu-id="31fe9-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31fe9-899">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31fe9-899">Returns:</span></span>

<span data-ttu-id="31fe9-900">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="31fe9-900">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="31fe9-901">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="31fe9-901">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="31fe9-902">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-902">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="31fe9-903">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="31fe9-903">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="31fe9-904">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="31fe9-904">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="31fe9-p158">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p158">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-908">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-908">Parameters</span></span>

|<span data-ttu-id="31fe9-909">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-909">Name</span></span>| <span data-ttu-id="31fe9-910">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-910">Type</span></span>| <span data-ttu-id="31fe9-911">Atributos</span><span class="sxs-lookup"><span data-stu-id="31fe9-911">Attributes</span></span>| <span data-ttu-id="31fe9-912">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-912">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="31fe9-913">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-913">function</span></span>||<span data-ttu-id="31fe9-914">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-914">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="31fe9-915">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31fe9-915">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="31fe9-916">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="31fe9-916">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="31fe9-917">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-917">Object</span></span>| <span data-ttu-id="31fe9-918">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-918">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-919">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-919">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="31fe9-920">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-920">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31fe9-921">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-921">Requirements</span></span>

|<span data-ttu-id="31fe9-922">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-922">Requirement</span></span>| <span data-ttu-id="31fe9-923">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-923">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-924">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-924">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-925">1.0</span><span class="sxs-lookup"><span data-stu-id="31fe9-925">1.0</span></span>|
|[<span data-ttu-id="31fe9-926">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-926">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-927">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-927">ReadItem</span></span>|
|[<span data-ttu-id="31fe9-928">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="31fe9-928">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-929">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="31fe9-929">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-930">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-930">Example</span></span>

<span data-ttu-id="31fe9-p161">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="31fe9-p161">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="31fe9-934">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="31fe9-934">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="31fe9-935">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="31fe9-935">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="31fe9-936">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="31fe9-936">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="31fe9-937">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="31fe9-937">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="31fe9-938">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="31fe9-938">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="31fe9-939">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-939">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31fe9-940">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="31fe9-940">Parameters</span></span>

|<span data-ttu-id="31fe9-941">Nome</span><span class="sxs-lookup"><span data-stu-id="31fe9-941">Name</span></span>| <span data-ttu-id="31fe9-942">Tipo</span><span class="sxs-lookup"><span data-stu-id="31fe9-942">Type</span></span>| <span data-ttu-id="31fe9-943">Atributos</span><span class="sxs-lookup"><span data-stu-id="31fe9-943">Attributes</span></span>| <span data-ttu-id="31fe9-944">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-944">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="31fe9-945">String</span><span class="sxs-lookup"><span data-stu-id="31fe9-945">String</span></span>||<span data-ttu-id="31fe9-946">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="31fe9-946">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="31fe9-947">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-947">Object</span></span>| <span data-ttu-id="31fe9-948">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-948">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-949">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="31fe9-949">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="31fe9-950">Objeto</span><span class="sxs-lookup"><span data-stu-id="31fe9-950">Object</span></span>| <span data-ttu-id="31fe9-951">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-951">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-952">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31fe9-952">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="31fe9-953">function</span><span class="sxs-lookup"><span data-stu-id="31fe9-953">function</span></span>| <span data-ttu-id="31fe9-954">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31fe9-954">&lt;optional&gt;</span></span>|<span data-ttu-id="31fe9-955">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31fe9-955">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="31fe9-956">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="31fe9-956">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="31fe9-957">Erros</span><span class="sxs-lookup"><span data-stu-id="31fe9-957">Errors</span></span>

| <span data-ttu-id="31fe9-958">Código de erro</span><span class="sxs-lookup"><span data-stu-id="31fe9-958">Error code</span></span> | <span data-ttu-id="31fe9-959">Descrição</span><span class="sxs-lookup"><span data-stu-id="31fe9-959">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="31fe9-960">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="31fe9-960">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31fe9-961">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31fe9-961">Requirements</span></span>

|<span data-ttu-id="31fe9-962">Requisito</span><span class="sxs-lookup"><span data-stu-id="31fe9-962">Requirement</span></span>| <span data-ttu-id="31fe9-963">Valor</span><span class="sxs-lookup"><span data-stu-id="31fe9-963">Value</span></span>|
|---|---|
|[<span data-ttu-id="31fe9-964">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31fe9-964">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31fe9-965">1.1</span><span class="sxs-lookup"><span data-stu-id="31fe9-965">1.1</span></span>|
|[<span data-ttu-id="31fe9-966">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31fe9-966">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31fe9-967">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="31fe9-967">ReadWriteItem</span></span>|
|[<span data-ttu-id="31fe9-968">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31fe9-968">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31fe9-969">Escrever</span><span class="sxs-lookup"><span data-stu-id="31fe9-969">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="31fe9-970">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31fe9-970">Example</span></span>

<span data-ttu-id="31fe9-971">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="31fe9-971">The following code removes an attachment with an identifier of '0'.</span></span>

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
