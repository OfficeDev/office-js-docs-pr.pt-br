---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: d3242f2bdabf464c262fdb8e6efd8695dc7ee330
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268499"
---
# <a name="item"></a><span data-ttu-id="39f34-102">item</span><span class="sxs-lookup"><span data-stu-id="39f34-102">item</span></span>

### <span data-ttu-id="39f34-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="39f34-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="39f34-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="39f34-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-107">Requirements</span></span>

|<span data-ttu-id="39f34-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-108">Requirement</span></span>| <span data-ttu-id="39f34-109">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-111">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-111">1.0</span></span>|
|[<span data-ttu-id="39f34-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="39f34-113">Restricted</span></span>|
|[<span data-ttu-id="39f34-114">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-115">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="39f34-116">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="39f34-116">Members and methods</span></span>

| <span data-ttu-id="39f34-117">Membro	</span><span class="sxs-lookup"><span data-stu-id="39f34-117">Member</span></span> | <span data-ttu-id="39f34-118">Tipo	</span><span class="sxs-lookup"><span data-stu-id="39f34-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="39f34-119">attachments</span><span class="sxs-lookup"><span data-stu-id="39f34-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="39f34-120">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-120">Member</span></span> |
| [<span data-ttu-id="39f34-121">bcc</span><span class="sxs-lookup"><span data-stu-id="39f34-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="39f34-122">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-122">Member</span></span> |
| [<span data-ttu-id="39f34-123">body</span><span class="sxs-lookup"><span data-stu-id="39f34-123">body</span></span>](#body-body) | <span data-ttu-id="39f34-124">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-124">Member</span></span> |
| [<span data-ttu-id="39f34-125">cc</span><span class="sxs-lookup"><span data-stu-id="39f34-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="39f34-126">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-126">Member</span></span> |
| [<span data-ttu-id="39f34-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="39f34-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="39f34-128">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-128">Member</span></span> |
| [<span data-ttu-id="39f34-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="39f34-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="39f34-130">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-130">Member</span></span> |
| [<span data-ttu-id="39f34-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="39f34-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="39f34-132">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-132">Member</span></span> |
| [<span data-ttu-id="39f34-133">end</span><span class="sxs-lookup"><span data-stu-id="39f34-133">end</span></span>](#end-datetime) | <span data-ttu-id="39f34-134">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-134">Member</span></span> |
| [<span data-ttu-id="39f34-135">from</span><span class="sxs-lookup"><span data-stu-id="39f34-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="39f34-136">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-136">Member</span></span> |
| [<span data-ttu-id="39f34-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="39f34-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="39f34-138">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-138">Member</span></span> |
| [<span data-ttu-id="39f34-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="39f34-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="39f34-140">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-140">Member</span></span> |
| [<span data-ttu-id="39f34-141">itemId</span><span class="sxs-lookup"><span data-stu-id="39f34-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="39f34-142">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-142">Member</span></span> |
| [<span data-ttu-id="39f34-143">itemType</span><span class="sxs-lookup"><span data-stu-id="39f34-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="39f34-144">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-144">Member</span></span> |
| [<span data-ttu-id="39f34-145">location</span><span class="sxs-lookup"><span data-stu-id="39f34-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="39f34-146">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-146">Member</span></span> |
| [<span data-ttu-id="39f34-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="39f34-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="39f34-148">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-148">Member</span></span> |
| [<span data-ttu-id="39f34-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="39f34-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="39f34-150">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-150">Member</span></span> |
| [<span data-ttu-id="39f34-151">organizer</span><span class="sxs-lookup"><span data-stu-id="39f34-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="39f34-152">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-152">Member</span></span> |
| [<span data-ttu-id="39f34-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="39f34-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="39f34-154">Member</span><span class="sxs-lookup"><span data-stu-id="39f34-154">Member</span></span> |
| [<span data-ttu-id="39f34-155">sender</span><span class="sxs-lookup"><span data-stu-id="39f34-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="39f34-156">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-156">Member</span></span> |
| [<span data-ttu-id="39f34-157">start</span><span class="sxs-lookup"><span data-stu-id="39f34-157">start</span></span>](#start-datetime) | <span data-ttu-id="39f34-158">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-158">Member</span></span> |
| [<span data-ttu-id="39f34-159">subject</span><span class="sxs-lookup"><span data-stu-id="39f34-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="39f34-160">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-160">Member</span></span> |
| [<span data-ttu-id="39f34-161">to</span><span class="sxs-lookup"><span data-stu-id="39f34-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="39f34-162">Membro</span><span class="sxs-lookup"><span data-stu-id="39f34-162">Member</span></span> |
| [<span data-ttu-id="39f34-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="39f34-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="39f34-164">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-164">Method</span></span> |
| [<span data-ttu-id="39f34-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="39f34-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="39f34-166">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-166">Method</span></span> |
| [<span data-ttu-id="39f34-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="39f34-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="39f34-168">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-168">Method</span></span> |
| [<span data-ttu-id="39f34-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="39f34-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="39f34-170">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-170">Method</span></span> |
| [<span data-ttu-id="39f34-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="39f34-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="39f34-172">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-172">Method</span></span> |
| [<span data-ttu-id="39f34-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="39f34-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="39f34-174">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-174">Method</span></span> |
| [<span data-ttu-id="39f34-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="39f34-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="39f34-176">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-176">Method</span></span> |
| [<span data-ttu-id="39f34-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="39f34-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="39f34-178">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-178">Method</span></span> |
| [<span data-ttu-id="39f34-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="39f34-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="39f34-180">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-180">Method</span></span> |
| [<span data-ttu-id="39f34-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="39f34-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="39f34-182">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-182">Method</span></span> |
| [<span data-ttu-id="39f34-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="39f34-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="39f34-184">Método</span><span class="sxs-lookup"><span data-stu-id="39f34-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="39f34-185">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-185">Example</span></span>

<span data-ttu-id="39f34-186">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="39f34-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="39f34-187">Membros</span><span class="sxs-lookup"><span data-stu-id="39f34-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="39f34-188">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="39f34-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="39f34-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-191">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="39f34-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="39f34-192">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="39f34-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-193">Type</span></span>

*   <span data-ttu-id="39f34-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="39f34-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-195">Requirements</span></span>

|<span data-ttu-id="39f34-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-196">Requirement</span></span>| <span data-ttu-id="39f34-197">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-199">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-199">1.0</span></span>|
|[<span data-ttu-id="39f34-200">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-201">ReadItem</span></span>|
|[<span data-ttu-id="39f34-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-203">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-204">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-204">Example</span></span>

<span data-ttu-id="39f34-205">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="39f34-206">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-207">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="39f34-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="39f34-208">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="39f34-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-209">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-209">Type</span></span>

*   [<span data-ttu-id="39f34-210">Destinatários</span><span class="sxs-lookup"><span data-stu-id="39f34-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-211">Requirements</span></span>

|<span data-ttu-id="39f34-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-212">Requirement</span></span>| <span data-ttu-id="39f34-213">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-215">1.1</span><span class="sxs-lookup"><span data-stu-id="39f34-215">1.1</span></span>|
|[<span data-ttu-id="39f34-216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-217">ReadItem</span></span>|
|[<span data-ttu-id="39f34-218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-219">Escrever</span><span class="sxs-lookup"><span data-stu-id="39f34-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-220">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-220">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="39f34-221">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-222">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="39f34-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-223">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-223">Type</span></span>

*   [<span data-ttu-id="39f34-224">Body</span><span class="sxs-lookup"><span data-stu-id="39f34-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-225">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-225">Requirements</span></span>

|<span data-ttu-id="39f34-226">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-226">Requirement</span></span>| <span data-ttu-id="39f34-227">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-228">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-229">1.1</span><span class="sxs-lookup"><span data-stu-id="39f34-229">1.1</span></span>|
|[<span data-ttu-id="39f34-230">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-231">ReadItem</span></span>|
|[<span data-ttu-id="39f34-232">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-233">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-234">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-234">Example</span></span>

<span data-ttu-id="39f34-235">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="39f34-235">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="39f34-236">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-236">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="39f34-237">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="39f34-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-238">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="39f34-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="39f34-239">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-240">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-240">Read mode</span></span>

<span data-ttu-id="39f34-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="39f34-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-243">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-243">Compose mode</span></span>

<span data-ttu-id="39f34-244">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="39f34-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="39f34-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-245">Type</span></span>

*   <span data-ttu-id="39f34-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-247">Requirements</span></span>

|<span data-ttu-id="39f34-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-248">Requirement</span></span>| <span data-ttu-id="39f34-249">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-251">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-251">1.0</span></span>|
|[<span data-ttu-id="39f34-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-253">ReadItem</span></span>|
|[<span data-ttu-id="39f34-254">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-255">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="39f34-256">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="39f34-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="39f34-257">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="39f34-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="39f34-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="39f34-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="39f34-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="39f34-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-262">Type</span></span>

*   <span data-ttu-id="39f34-263">String</span><span class="sxs-lookup"><span data-stu-id="39f34-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-264">Requirements</span></span>

|<span data-ttu-id="39f34-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-265">Requirement</span></span>| <span data-ttu-id="39f34-266">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-268">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-268">1.0</span></span>|
|[<span data-ttu-id="39f34-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-270">ReadItem</span></span>|
|[<span data-ttu-id="39f34-271">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-273">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="39f34-274">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="39f34-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="39f34-p110">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-277">Type</span></span>

*   <span data-ttu-id="39f34-278">Data</span><span class="sxs-lookup"><span data-stu-id="39f34-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-279">Requirements</span></span>

|<span data-ttu-id="39f34-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-280">Requirement</span></span>| <span data-ttu-id="39f34-281">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-283">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-283">1.0</span></span>|
|[<span data-ttu-id="39f34-284">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-285">ReadItem</span></span>|
|[<span data-ttu-id="39f34-286">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-287">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-288">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="39f34-289">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="39f34-289">dateTimeModified: Date</span></span>

<span data-ttu-id="39f34-290">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="39f34-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="39f34-291">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-292">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-293">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-293">Type</span></span>

*   <span data-ttu-id="39f34-294">Data</span><span class="sxs-lookup"><span data-stu-id="39f34-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-295">Requirements</span></span>

|<span data-ttu-id="39f34-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-296">Requirement</span></span>| <span data-ttu-id="39f34-297">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-299">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-299">1.0</span></span>|
|[<span data-ttu-id="39f34-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-301">ReadItem</span></span>|
|[<span data-ttu-id="39f34-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-303">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-304">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-304">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="39f34-305">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-306">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="39f34-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="39f34-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="39f34-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-309">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-309">Read mode</span></span>

<span data-ttu-id="39f34-310">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="39f34-310">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-311">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-311">Compose mode</span></span>

<span data-ttu-id="39f34-312">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="39f34-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="39f34-313">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="39f34-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="39f34-314">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="39f34-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="39f34-315">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-315">Type</span></span>

*   <span data-ttu-id="39f34-316">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-317">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-317">Requirements</span></span>

|<span data-ttu-id="39f34-318">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-318">Requirement</span></span>| <span data-ttu-id="39f34-319">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-320">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-321">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-321">1.0</span></span>|
|[<span data-ttu-id="39f34-322">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-323">ReadItem</span></span>|
|[<span data-ttu-id="39f34-324">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-325">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-325">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="39f34-326">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="39f34-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="39f34-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-331">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="39f34-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-332">Type</span></span>

*   [<span data-ttu-id="39f34-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="39f34-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-334">Requirements</span></span>

|<span data-ttu-id="39f34-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-335">Requirement</span></span>| <span data-ttu-id="39f34-336">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-337">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-338">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-338">1.0</span></span>|
|[<span data-ttu-id="39f34-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-340">ReadItem</span></span>|
|[<span data-ttu-id="39f34-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-342">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-343">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="39f34-344">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="39f34-344">internetMessageId: String</span></span>

<span data-ttu-id="39f34-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-347">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-347">Type</span></span>

*   <span data-ttu-id="39f34-348">String</span><span class="sxs-lookup"><span data-stu-id="39f34-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-349">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-349">Requirements</span></span>

|<span data-ttu-id="39f34-350">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-350">Requirement</span></span>| <span data-ttu-id="39f34-351">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-352">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-353">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-353">1.0</span></span>|
|[<span data-ttu-id="39f34-354">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-355">ReadItem</span></span>|
|[<span data-ttu-id="39f34-356">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-357">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-358">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-358">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="39f34-359">doclass: String</span><span class="sxs-lookup"><span data-stu-id="39f34-359">itemClass: String</span></span>

<span data-ttu-id="39f34-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="39f34-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="39f34-364">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-364">Type</span></span> | <span data-ttu-id="39f34-365">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-365">Description</span></span> | <span data-ttu-id="39f34-366">classe de item</span><span class="sxs-lookup"><span data-stu-id="39f34-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="39f34-367">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="39f34-367">Appointment items</span></span> | <span data-ttu-id="39f34-368">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="39f34-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="39f34-369">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="39f34-369">Message items</span></span> | <span data-ttu-id="39f34-370">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="39f34-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="39f34-371">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="39f34-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-372">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-372">Type</span></span>

*   <span data-ttu-id="39f34-373">String</span><span class="sxs-lookup"><span data-stu-id="39f34-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-374">Requirements</span></span>

|<span data-ttu-id="39f34-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-375">Requirement</span></span>| <span data-ttu-id="39f34-376">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-378">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-378">1.0</span></span>|
|[<span data-ttu-id="39f34-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-380">ReadItem</span></span>|
|[<span data-ttu-id="39f34-381">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-382">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-383">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="39f34-384">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="39f34-384">(nullable) itemId: String</span></span>

<span data-ttu-id="39f34-385">Obtém o identificador do item dos Serviços Web do Exchange para o item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="39f34-386">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-387">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="39f34-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="39f34-388">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="39f34-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="39f34-389">Antes de fazer chamadas da API REST usando esse valor, ele deve ser `Office.context.mailbox.convertToRestId`convertido usando o, que está disponível a partir do conjunto de requisitos 1,3.</span><span class="sxs-lookup"><span data-stu-id="39f34-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="39f34-390">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="39f34-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-391">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-391">Type</span></span>

*   <span data-ttu-id="39f34-392">String</span><span class="sxs-lookup"><span data-stu-id="39f34-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-393">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-393">Requirements</span></span>

|<span data-ttu-id="39f34-394">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-394">Requirement</span></span>| <span data-ttu-id="39f34-395">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-396">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-397">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-397">1.0</span></span>|
|[<span data-ttu-id="39f34-398">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-399">ReadItem</span></span>|
|[<span data-ttu-id="39f34-400">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-401">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-402">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-402">Example</span></span>

<span data-ttu-id="39f34-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="39f34-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="39f34-405">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-406">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="39f34-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="39f34-407">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-408">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-408">Type</span></span>

*   [<span data-ttu-id="39f34-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="39f34-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-410">Requirements</span></span>

|<span data-ttu-id="39f34-411">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-411">Requirement</span></span>| <span data-ttu-id="39f34-412">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-414">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-414">1.0</span></span>|
|[<span data-ttu-id="39f34-415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-416">ReadItem</span></span>|
|[<span data-ttu-id="39f34-417">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-418">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-419">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="39f34-420">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-421">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-422">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-422">Read mode</span></span>

<span data-ttu-id="39f34-423">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-424">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-424">Compose mode</span></span>

<span data-ttu-id="39f34-425">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="39f34-426">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-426">Type</span></span>

*   <span data-ttu-id="39f34-427">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-428">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-428">Requirements</span></span>

|<span data-ttu-id="39f34-429">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-429">Requirement</span></span>| <span data-ttu-id="39f34-430">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-431">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-432">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-432">1.0</span></span>|
|[<span data-ttu-id="39f34-433">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-434">ReadItem</span></span>|
|[<span data-ttu-id="39f34-435">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-436">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-436">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="39f34-437">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="39f34-437">normalizedSubject: String</span></span>

<span data-ttu-id="39f34-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="39f34-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="39f34-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-442">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-442">Type</span></span>

*   <span data-ttu-id="39f34-443">String</span><span class="sxs-lookup"><span data-stu-id="39f34-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-444">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-444">Requirements</span></span>

|<span data-ttu-id="39f34-445">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-445">Requirement</span></span>| <span data-ttu-id="39f34-446">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-447">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-448">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-448">1.0</span></span>|
|[<span data-ttu-id="39f34-449">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-450">ReadItem</span></span>|
|[<span data-ttu-id="39f34-451">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-452">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-453">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-453">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="39f34-454">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="39f34-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-455">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="39f34-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="39f34-456">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-457">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-457">Read mode</span></span>

<span data-ttu-id="39f34-458">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="39f34-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-459">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-459">Compose mode</span></span>

<span data-ttu-id="39f34-460">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="39f34-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="39f34-461">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-461">Type</span></span>

*   <span data-ttu-id="39f34-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-463">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-463">Requirements</span></span>

|<span data-ttu-id="39f34-464">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-464">Requirement</span></span>| <span data-ttu-id="39f34-465">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-466">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-467">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-467">1.0</span></span>|
|[<span data-ttu-id="39f34-468">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-469">ReadItem</span></span>|
|[<span data-ttu-id="39f34-470">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-471">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-471">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="39f34-472">organizador: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-475">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-475">Type</span></span>

*   [<span data-ttu-id="39f34-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="39f34-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-477">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-477">Requirements</span></span>

|<span data-ttu-id="39f34-478">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-478">Requirement</span></span>| <span data-ttu-id="39f34-479">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-480">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-481">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-481">1.0</span></span>|
|[<span data-ttu-id="39f34-482">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-483">ReadItem</span></span>|
|[<span data-ttu-id="39f34-484">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-485">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-486">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-486">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="39f34-487">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="39f34-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-488">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="39f34-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="39f34-489">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-490">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-490">Read mode</span></span>

<span data-ttu-id="39f34-491">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="39f34-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-492">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-492">Compose mode</span></span>

<span data-ttu-id="39f34-493">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="39f34-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="39f34-494">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-494">Type</span></span>

*   <span data-ttu-id="39f34-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-496">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-496">Requirements</span></span>

|<span data-ttu-id="39f34-497">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-497">Requirement</span></span>| <span data-ttu-id="39f34-498">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-499">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-500">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-500">1.0</span></span>|
|[<span data-ttu-id="39f34-501">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-502">ReadItem</span></span>|
|[<span data-ttu-id="39f34-503">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-504">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-504">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="39f34-505">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="39f34-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="39f34-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="39f34-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-510">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="39f34-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="39f34-511">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-511">Type</span></span>

*   [<span data-ttu-id="39f34-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="39f34-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="39f34-513">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-513">Requirements</span></span>

|<span data-ttu-id="39f34-514">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-514">Requirement</span></span>| <span data-ttu-id="39f34-515">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-516">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-517">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-517">1.0</span></span>|
|[<span data-ttu-id="39f34-518">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-519">ReadItem</span></span>|
|[<span data-ttu-id="39f34-520">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-521">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-522">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-522">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="39f34-523">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-524">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="39f34-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="39f34-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="39f34-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-527">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-527">Read mode</span></span>

<span data-ttu-id="39f34-528">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="39f34-528">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-529">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-529">Compose mode</span></span>

<span data-ttu-id="39f34-530">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="39f34-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="39f34-531">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="39f34-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="39f34-532">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="39f34-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="39f34-533">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-533">Type</span></span>

*   <span data-ttu-id="39f34-534">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-535">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-535">Requirements</span></span>

|<span data-ttu-id="39f34-536">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-536">Requirement</span></span>| <span data-ttu-id="39f34-537">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-538">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-539">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-539">1.0</span></span>|
|[<span data-ttu-id="39f34-540">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-541">ReadItem</span></span>|
|[<span data-ttu-id="39f34-542">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-543">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-543">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="39f34-544">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-545">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="39f34-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="39f34-546">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="39f34-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-547">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-547">Read mode</span></span>

<span data-ttu-id="39f34-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="39f34-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-550">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-550">Compose mode</span></span>

<span data-ttu-id="39f34-551">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="39f34-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="39f34-552">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-552">Type</span></span>

*   <span data-ttu-id="39f34-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-554">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-554">Requirements</span></span>

|<span data-ttu-id="39f34-555">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-555">Requirement</span></span>| <span data-ttu-id="39f34-556">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-557">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-558">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-558">1.0</span></span>|
|[<span data-ttu-id="39f34-559">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-560">ReadItem</span></span>|
|[<span data-ttu-id="39f34-561">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-562">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-562">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="39f34-563">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="39f34-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="39f34-564">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="39f34-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="39f34-565">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="39f34-566">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="39f34-566">Read mode</span></span>

<span data-ttu-id="39f34-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="39f34-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="39f34-569">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="39f34-569">Compose mode</span></span>

<span data-ttu-id="39f34-570">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="39f34-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="39f34-571">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-571">Type</span></span>

*   <span data-ttu-id="39f34-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-573">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-573">Requirements</span></span>

|<span data-ttu-id="39f34-574">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-574">Requirement</span></span>| <span data-ttu-id="39f34-575">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-576">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-577">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-577">1.0</span></span>|
|[<span data-ttu-id="39f34-578">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-579">ReadItem</span></span>|
|[<span data-ttu-id="39f34-580">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-581">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="39f34-582">Métodos</span><span class="sxs-lookup"><span data-stu-id="39f34-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="39f34-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="39f34-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="39f34-584">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="39f34-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="39f34-585">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="39f34-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="39f34-586">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="39f34-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-587">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-587">Parameters</span></span>

|<span data-ttu-id="39f34-588">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-588">Name</span></span>| <span data-ttu-id="39f34-589">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-589">Type</span></span>| <span data-ttu-id="39f34-590">Atributos</span><span class="sxs-lookup"><span data-stu-id="39f34-590">Attributes</span></span>| <span data-ttu-id="39f34-591">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="39f34-592">String</span><span class="sxs-lookup"><span data-stu-id="39f34-592">String</span></span>||<span data-ttu-id="39f34-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="39f34-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="39f34-595">String</span><span class="sxs-lookup"><span data-stu-id="39f34-595">String</span></span>||<span data-ttu-id="39f34-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="39f34-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="39f34-598">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-598">Object</span></span>| <span data-ttu-id="39f34-599">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-599">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-600">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="39f34-601">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-601">Object</span></span>| <span data-ttu-id="39f34-602">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-602">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-603">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="39f34-604">function</span><span class="sxs-lookup"><span data-stu-id="39f34-604">function</span></span>| <span data-ttu-id="39f34-605">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-605">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-606">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="39f34-607">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39f34-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="39f34-608">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="39f34-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="39f34-609">Erros</span><span class="sxs-lookup"><span data-stu-id="39f34-609">Errors</span></span>

| <span data-ttu-id="39f34-610">Código de erro</span><span class="sxs-lookup"><span data-stu-id="39f34-610">Error code</span></span> | <span data-ttu-id="39f34-611">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="39f34-612">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="39f34-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="39f34-613">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="39f34-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="39f34-614">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="39f34-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39f34-615">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-615">Requirements</span></span>

|<span data-ttu-id="39f34-616">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-616">Requirement</span></span>| <span data-ttu-id="39f34-617">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-618">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-619">1.1</span><span class="sxs-lookup"><span data-stu-id="39f34-619">1.1</span></span>|
|[<span data-ttu-id="39f34-620">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="39f34-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="39f34-622">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-623">Escrever</span><span class="sxs-lookup"><span data-stu-id="39f34-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-624">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="39f34-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="39f34-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="39f34-626">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="39f34-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="39f34-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="39f34-630">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="39f34-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="39f34-631">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="39f34-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-632">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-632">Parameters</span></span>

|<span data-ttu-id="39f34-633">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-633">Name</span></span>| <span data-ttu-id="39f34-634">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-634">Type</span></span>| <span data-ttu-id="39f34-635">Atributos</span><span class="sxs-lookup"><span data-stu-id="39f34-635">Attributes</span></span>| <span data-ttu-id="39f34-636">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="39f34-637">String</span><span class="sxs-lookup"><span data-stu-id="39f34-637">String</span></span>||<span data-ttu-id="39f34-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="39f34-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="39f34-640">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="39f34-640">String</span></span>||<span data-ttu-id="39f34-641">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="39f34-641">The subject of the item to be attached.</span></span> <span data-ttu-id="39f34-642">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="39f34-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="39f34-643">Object</span><span class="sxs-lookup"><span data-stu-id="39f34-643">Object</span></span>| <span data-ttu-id="39f34-644">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-644">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-645">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="39f34-646">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-646">Object</span></span>| <span data-ttu-id="39f34-647">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-647">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-648">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="39f34-649">function</span><span class="sxs-lookup"><span data-stu-id="39f34-649">function</span></span>| <span data-ttu-id="39f34-650">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-650">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-651">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="39f34-652">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39f34-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="39f34-653">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="39f34-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="39f34-654">Erros</span><span class="sxs-lookup"><span data-stu-id="39f34-654">Errors</span></span>

| <span data-ttu-id="39f34-655">Código de erro</span><span class="sxs-lookup"><span data-stu-id="39f34-655">Error code</span></span> | <span data-ttu-id="39f34-656">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="39f34-657">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="39f34-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39f34-658">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-658">Requirements</span></span>

|<span data-ttu-id="39f34-659">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-659">Requirement</span></span>| <span data-ttu-id="39f34-660">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-661">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-662">1.1</span><span class="sxs-lookup"><span data-stu-id="39f34-662">1.1</span></span>|
|[<span data-ttu-id="39f34-663">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="39f34-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="39f34-665">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-666">Escrever</span><span class="sxs-lookup"><span data-stu-id="39f34-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-667">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-667">Example</span></span>

<span data-ttu-id="39f34-668">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="39f34-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="39f34-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="39f34-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="39f34-670">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="39f34-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-671">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="39f34-672">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="39f34-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="39f34-673">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="39f34-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-674">A capacidade de incluir anexos na chamada para `displayReplyAllForm` não é suportada no conjunto de requisitos 1,1.</span><span class="sxs-lookup"><span data-stu-id="39f34-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="39f34-675">O suporte a anexos foi adicionado a `displayReplyAllForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="39f34-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-676">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-676">Parameters</span></span>

|<span data-ttu-id="39f34-677">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-677">Name</span></span>| <span data-ttu-id="39f34-678">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-678">Type</span></span>| <span data-ttu-id="39f34-679">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="39f34-680">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="39f34-680">String &#124; Object</span></span>| |<span data-ttu-id="39f34-p138">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="39f34-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="39f34-683">**OU**</span><span class="sxs-lookup"><span data-stu-id="39f34-683">**OR**</span></span><br/><span data-ttu-id="39f34-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="39f34-686">String</span><span class="sxs-lookup"><span data-stu-id="39f34-686">String</span></span> | <span data-ttu-id="39f34-687">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-687">&lt;optional&gt;</span></span> | <span data-ttu-id="39f34-p140">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="39f34-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="39f34-690">function</span><span class="sxs-lookup"><span data-stu-id="39f34-690">function</span></span> | <span data-ttu-id="39f34-691">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-691">&lt;optional&gt;</span></span> | <span data-ttu-id="39f34-692">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39f34-693">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-693">Requirements</span></span>

|<span data-ttu-id="39f34-694">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-694">Requirement</span></span>| <span data-ttu-id="39f34-695">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-696">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-697">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-697">1.0</span></span>|
|[<span data-ttu-id="39f34-698">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-699">ReadItem</span></span>|
|[<span data-ttu-id="39f34-700">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-701">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="39f34-702">Exemplos</span><span class="sxs-lookup"><span data-stu-id="39f34-702">Examples</span></span>

<span data-ttu-id="39f34-703">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="39f34-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="39f34-704">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="39f34-704">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="39f34-705">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="39f34-705">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="39f34-706">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-706">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="39f34-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="39f34-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="39f34-708">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="39f34-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-709">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="39f34-710">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="39f34-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="39f34-711">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="39f34-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-712">A capacidade de incluir anexos na chamada para `displayReplyForm` não é suportada no conjunto de requisitos 1,1.</span><span class="sxs-lookup"><span data-stu-id="39f34-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="39f34-713">O suporte a anexos foi adicionado a `displayReplyForm` no conjunto de requisitos 1.2 e acima.</span><span class="sxs-lookup"><span data-stu-id="39f34-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-714">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-714">Parameters</span></span>

|<span data-ttu-id="39f34-715">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-715">Name</span></span>| <span data-ttu-id="39f34-716">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-716">Type</span></span>| <span data-ttu-id="39f34-717">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="39f34-718">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="39f34-718">String &#124; Object</span></span>| | <span data-ttu-id="39f34-p142">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="39f34-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="39f34-721">**OU**</span><span class="sxs-lookup"><span data-stu-id="39f34-721">**OR**</span></span><br/><span data-ttu-id="39f34-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="39f34-724">String</span><span class="sxs-lookup"><span data-stu-id="39f34-724">String</span></span> | <span data-ttu-id="39f34-725">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-725">&lt;optional&gt;</span></span> | <span data-ttu-id="39f34-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="39f34-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="39f34-728">function</span><span class="sxs-lookup"><span data-stu-id="39f34-728">function</span></span> | <span data-ttu-id="39f34-729">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-729">&lt;optional&gt;</span></span> | <span data-ttu-id="39f34-730">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39f34-731">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-731">Requirements</span></span>

|<span data-ttu-id="39f34-732">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-732">Requirement</span></span>| <span data-ttu-id="39f34-733">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-734">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-735">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-735">1.0</span></span>|
|[<span data-ttu-id="39f34-736">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-737">ReadItem</span></span>|
|[<span data-ttu-id="39f34-738">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-739">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="39f34-740">Exemplos</span><span class="sxs-lookup"><span data-stu-id="39f34-740">Examples</span></span>

<span data-ttu-id="39f34-741">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="39f34-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="39f34-742">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="39f34-742">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="39f34-743">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="39f34-743">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="39f34-744">Responder com um corpo e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-744">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="39f34-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="39f34-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="39f34-746">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="39f34-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-747">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-748">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-748">Requirements</span></span>

|<span data-ttu-id="39f34-749">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-749">Requirement</span></span>| <span data-ttu-id="39f34-750">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-751">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-752">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-752">1.0</span></span>|
|[<span data-ttu-id="39f34-753">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-754">ReadItem</span></span>|
|[<span data-ttu-id="39f34-755">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-756">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39f34-757">Retorna:</span><span class="sxs-lookup"><span data-stu-id="39f34-757">Returns:</span></span>

<span data-ttu-id="39f34-758">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="39f34-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="39f34-759">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-759">Example</span></span>

<span data-ttu-id="39f34-760">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-760">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="39f34-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="39f34-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="39f34-762">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="39f34-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-763">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-764">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-764">Parameters</span></span>

|<span data-ttu-id="39f34-765">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-765">Name</span></span>| <span data-ttu-id="39f34-766">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-766">Type</span></span>| <span data-ttu-id="39f34-767">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="39f34-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="39f34-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="39f34-769">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="39f34-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39f34-770">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-770">Requirements</span></span>

|<span data-ttu-id="39f34-771">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-771">Requirement</span></span>| <span data-ttu-id="39f34-772">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-773">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-774">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-774">1.0</span></span>|
|[<span data-ttu-id="39f34-775">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-776">Restrito</span><span class="sxs-lookup"><span data-stu-id="39f34-776">Restricted</span></span>|
|[<span data-ttu-id="39f34-777">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-778">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39f34-779">Retorna:</span><span class="sxs-lookup"><span data-stu-id="39f34-779">Returns:</span></span>

<span data-ttu-id="39f34-780">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="39f34-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="39f34-781">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="39f34-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="39f34-782">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="39f34-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="39f34-783">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="39f34-784">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="39f34-784">Value of `entityType`</span></span> | <span data-ttu-id="39f34-785">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="39f34-785">Type of objects in returned array</span></span> | <span data-ttu-id="39f34-786">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="39f34-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="39f34-787">String</span><span class="sxs-lookup"><span data-stu-id="39f34-787">String</span></span> | <span data-ttu-id="39f34-788">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="39f34-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="39f34-789">Contato</span><span class="sxs-lookup"><span data-stu-id="39f34-789">Contact</span></span> | <span data-ttu-id="39f34-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="39f34-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="39f34-791">String</span><span class="sxs-lookup"><span data-stu-id="39f34-791">String</span></span> | <span data-ttu-id="39f34-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="39f34-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="39f34-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="39f34-793">MeetingSuggestion</span></span> | <span data-ttu-id="39f34-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="39f34-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="39f34-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="39f34-795">PhoneNumber</span></span> | <span data-ttu-id="39f34-796">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="39f34-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="39f34-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="39f34-797">TaskSuggestion</span></span> | <span data-ttu-id="39f34-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="39f34-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="39f34-799">String</span><span class="sxs-lookup"><span data-stu-id="39f34-799">String</span></span> | <span data-ttu-id="39f34-800">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="39f34-800">**Restricted**</span></span> |

<span data-ttu-id="39f34-801">Tipo:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="39f34-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="39f34-802">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-802">Example</span></span>

<span data-ttu-id="39f34-803">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="39f34-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="39f34-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="39f34-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="39f34-805">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="39f34-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-806">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="39f34-807">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="39f34-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-808">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-808">Parameters</span></span>

|<span data-ttu-id="39f34-809">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-809">Name</span></span>| <span data-ttu-id="39f34-810">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-810">Type</span></span>| <span data-ttu-id="39f34-811">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="39f34-812">String</span><span class="sxs-lookup"><span data-stu-id="39f34-812">String</span></span>|<span data-ttu-id="39f34-813">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="39f34-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39f34-814">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-814">Requirements</span></span>

|<span data-ttu-id="39f34-815">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-815">Requirement</span></span>| <span data-ttu-id="39f34-816">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-817">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-818">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-818">1.0</span></span>|
|[<span data-ttu-id="39f34-819">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-820">ReadItem</span></span>|
|[<span data-ttu-id="39f34-821">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-822">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39f34-823">Retorna:</span><span class="sxs-lookup"><span data-stu-id="39f34-823">Returns:</span></span>

<span data-ttu-id="39f34-p146">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="39f34-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="39f34-826">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="39f34-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="39f34-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="39f34-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="39f34-828">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="39f34-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-829">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="39f34-p147">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="39f34-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="39f34-833">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="39f34-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="39f34-834">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="39f34-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="39f34-p148">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="39f34-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f34-837">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-837">Requirements</span></span>

|<span data-ttu-id="39f34-838">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-838">Requirement</span></span>| <span data-ttu-id="39f34-839">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-840">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-841">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-841">1.0</span></span>|
|[<span data-ttu-id="39f34-842">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-843">ReadItem</span></span>|
|[<span data-ttu-id="39f34-844">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-845">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39f34-846">Retorna:</span><span class="sxs-lookup"><span data-stu-id="39f34-846">Returns:</span></span>

<span data-ttu-id="39f34-p149">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="39f34-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="39f34-849">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="39f34-849">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="39f34-850">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-850">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="39f34-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-851">Example</span></span>

<span data-ttu-id="39f34-852">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="39f34-852">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="39f34-853">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="39f34-853">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="39f34-854">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="39f34-854">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="39f34-855">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="39f34-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="39f34-856">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="39f34-856">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="39f34-p150">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="39f34-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-859">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-859">Parameters</span></span>

|<span data-ttu-id="39f34-860">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-860">Name</span></span>| <span data-ttu-id="39f34-861">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-861">Type</span></span>| <span data-ttu-id="39f34-862">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-862">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="39f34-863">String</span><span class="sxs-lookup"><span data-stu-id="39f34-863">String</span></span>|<span data-ttu-id="39f34-864">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="39f34-864">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39f34-865">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-865">Requirements</span></span>

|<span data-ttu-id="39f34-866">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-866">Requirement</span></span>| <span data-ttu-id="39f34-867">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-867">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-868">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-868">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-869">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-869">1.0</span></span>|
|[<span data-ttu-id="39f34-870">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-870">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-871">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-871">ReadItem</span></span>|
|[<span data-ttu-id="39f34-872">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-872">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-873">Read</span><span class="sxs-lookup"><span data-stu-id="39f34-873">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="39f34-874">Retorna:</span><span class="sxs-lookup"><span data-stu-id="39f34-874">Returns:</span></span>

<span data-ttu-id="39f34-875">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="39f34-875">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="39f34-876">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="39f34-876">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="39f34-877">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="39f34-877">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="39f34-878">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-878">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="39f34-879">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="39f34-879">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="39f34-880">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="39f34-880">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="39f34-p151">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="39f34-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-884">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-884">Parameters</span></span>

|<span data-ttu-id="39f34-885">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-885">Name</span></span>| <span data-ttu-id="39f34-886">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-886">Type</span></span>| <span data-ttu-id="39f34-887">Atributos</span><span class="sxs-lookup"><span data-stu-id="39f34-887">Attributes</span></span>| <span data-ttu-id="39f34-888">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-888">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="39f34-889">function</span><span class="sxs-lookup"><span data-stu-id="39f34-889">function</span></span>||<span data-ttu-id="39f34-890">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-890">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="39f34-891">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="39f34-891">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="39f34-892">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="39f34-892">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="39f34-893">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-893">Object</span></span>| <span data-ttu-id="39f34-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-894">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-895">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-895">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="39f34-896">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-896">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39f34-897">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-897">Requirements</span></span>

|<span data-ttu-id="39f34-898">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-898">Requirement</span></span>| <span data-ttu-id="39f34-899">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-900">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-901">1.0</span><span class="sxs-lookup"><span data-stu-id="39f34-901">1.0</span></span>|
|[<span data-ttu-id="39f34-902">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-902">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-903">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39f34-903">ReadItem</span></span>|
|[<span data-ttu-id="39f34-904">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="39f34-904">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-905">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="39f34-905">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-906">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-906">Example</span></span>

<span data-ttu-id="39f34-p154">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="39f34-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="39f34-910">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="39f34-910">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="39f34-911">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="39f34-911">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="39f34-912">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="39f34-912">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="39f34-913">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="39f34-913">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="39f34-914">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="39f34-914">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="39f34-915">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="39f34-915">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="39f34-916">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="39f34-916">Parameters</span></span>

|<span data-ttu-id="39f34-917">Nome</span><span class="sxs-lookup"><span data-stu-id="39f34-917">Name</span></span>| <span data-ttu-id="39f34-918">Tipo</span><span class="sxs-lookup"><span data-stu-id="39f34-918">Type</span></span>| <span data-ttu-id="39f34-919">Atributos</span><span class="sxs-lookup"><span data-stu-id="39f34-919">Attributes</span></span>| <span data-ttu-id="39f34-920">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-920">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="39f34-921">String</span><span class="sxs-lookup"><span data-stu-id="39f34-921">String</span></span>||<span data-ttu-id="39f34-922">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="39f34-922">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="39f34-923">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-923">Object</span></span>| <span data-ttu-id="39f34-924">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-924">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-925">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="39f34-925">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="39f34-926">Objeto</span><span class="sxs-lookup"><span data-stu-id="39f34-926">Object</span></span>| <span data-ttu-id="39f34-927">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-927">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-928">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="39f34-928">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="39f34-929">function</span><span class="sxs-lookup"><span data-stu-id="39f34-929">function</span></span>| <span data-ttu-id="39f34-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="39f34-930">&lt;optional&gt;</span></span>|<span data-ttu-id="39f34-931">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="39f34-931">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="39f34-932">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="39f34-932">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="39f34-933">Erros</span><span class="sxs-lookup"><span data-stu-id="39f34-933">Errors</span></span>

| <span data-ttu-id="39f34-934">Código de erro</span><span class="sxs-lookup"><span data-stu-id="39f34-934">Error code</span></span> | <span data-ttu-id="39f34-935">Descrição</span><span class="sxs-lookup"><span data-stu-id="39f34-935">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="39f34-936">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="39f34-936">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="39f34-937">Requisitos</span><span class="sxs-lookup"><span data-stu-id="39f34-937">Requirements</span></span>

|<span data-ttu-id="39f34-938">Requisito</span><span class="sxs-lookup"><span data-stu-id="39f34-938">Requirement</span></span>| <span data-ttu-id="39f34-939">Valor</span><span class="sxs-lookup"><span data-stu-id="39f34-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f34-940">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="39f34-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f34-941">1.1</span><span class="sxs-lookup"><span data-stu-id="39f34-941">1.1</span></span>|
|[<span data-ttu-id="39f34-942">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="39f34-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f34-943">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="39f34-943">ReadWriteItem</span></span>|
|[<span data-ttu-id="39f34-944">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="39f34-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39f34-945">Escrever</span><span class="sxs-lookup"><span data-stu-id="39f34-945">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="39f34-946">Exemplo</span><span class="sxs-lookup"><span data-stu-id="39f34-946">Example</span></span>

<span data-ttu-id="39f34-947">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="39f34-947">The following code removes an attachment with an identifier of '0'.</span></span>

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
