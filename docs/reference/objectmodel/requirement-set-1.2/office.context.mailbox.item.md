---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,2
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: c765b0901c15adb7c3651ac279f224de05002023
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167344"
---
# <a name="item"></a><span data-ttu-id="a0ec0-102">item</span><span class="sxs-lookup"><span data-stu-id="a0ec0-102">item</span></span>

### <span data-ttu-id="a0ec0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="a0ec0-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-107">Requirements</span></span>

|<span data-ttu-id="a0ec0-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-108">Requirement</span></span>| <span data-ttu-id="a0ec0-109">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-111">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-111">1.0</span></span>|
|[<span data-ttu-id="a0ec0-112">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-113">Restrito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-113">Restricted</span></span>|
|[<span data-ttu-id="a0ec0-114">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-115">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a0ec0-116">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-116">Members and methods</span></span>

| <span data-ttu-id="a0ec0-117">Membro	</span><span class="sxs-lookup"><span data-stu-id="a0ec0-117">Member</span></span> | <span data-ttu-id="a0ec0-118">Tipo	</span><span class="sxs-lookup"><span data-stu-id="a0ec0-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a0ec0-119">attachments</span><span class="sxs-lookup"><span data-stu-id="a0ec0-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="a0ec0-120">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-120">Member</span></span> |
| [<span data-ttu-id="a0ec0-121">bcc</span><span class="sxs-lookup"><span data-stu-id="a0ec0-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="a0ec0-122">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-122">Member</span></span> |
| [<span data-ttu-id="a0ec0-123">body</span><span class="sxs-lookup"><span data-stu-id="a0ec0-123">body</span></span>](#body-body) | <span data-ttu-id="a0ec0-124">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-124">Member</span></span> |
| [<span data-ttu-id="a0ec0-125">cc</span><span class="sxs-lookup"><span data-stu-id="a0ec0-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0ec0-126">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-126">Member</span></span> |
| [<span data-ttu-id="a0ec0-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="a0ec0-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a0ec0-128">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-128">Member</span></span> |
| [<span data-ttu-id="a0ec0-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a0ec0-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a0ec0-130">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-130">Member</span></span> |
| [<span data-ttu-id="a0ec0-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a0ec0-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a0ec0-132">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-132">Member</span></span> |
| [<span data-ttu-id="a0ec0-133">end</span><span class="sxs-lookup"><span data-stu-id="a0ec0-133">end</span></span>](#end-datetime) | <span data-ttu-id="a0ec0-134">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-134">Member</span></span> |
| [<span data-ttu-id="a0ec0-135">from</span><span class="sxs-lookup"><span data-stu-id="a0ec0-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="a0ec0-136">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-136">Member</span></span> |
| [<span data-ttu-id="a0ec0-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a0ec0-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a0ec0-138">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-138">Member</span></span> |
| [<span data-ttu-id="a0ec0-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="a0ec0-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a0ec0-140">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-140">Member</span></span> |
| [<span data-ttu-id="a0ec0-141">itemId</span><span class="sxs-lookup"><span data-stu-id="a0ec0-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a0ec0-142">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-142">Member</span></span> |
| [<span data-ttu-id="a0ec0-143">itemType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="a0ec0-144">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-144">Member</span></span> |
| [<span data-ttu-id="a0ec0-145">location</span><span class="sxs-lookup"><span data-stu-id="a0ec0-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="a0ec0-146">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-146">Member</span></span> |
| [<span data-ttu-id="a0ec0-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a0ec0-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a0ec0-148">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-148">Member</span></span> |
| [<span data-ttu-id="a0ec0-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a0ec0-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0ec0-150">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-150">Member</span></span> |
| [<span data-ttu-id="a0ec0-151">organizer</span><span class="sxs-lookup"><span data-stu-id="a0ec0-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="a0ec0-152">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-152">Member</span></span> |
| [<span data-ttu-id="a0ec0-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a0ec0-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0ec0-154">Member</span><span class="sxs-lookup"><span data-stu-id="a0ec0-154">Member</span></span> |
| [<span data-ttu-id="a0ec0-155">sender</span><span class="sxs-lookup"><span data-stu-id="a0ec0-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="a0ec0-156">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-156">Member</span></span> |
| [<span data-ttu-id="a0ec0-157">start</span><span class="sxs-lookup"><span data-stu-id="a0ec0-157">start</span></span>](#start-datetime) | <span data-ttu-id="a0ec0-158">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-158">Member</span></span> |
| [<span data-ttu-id="a0ec0-159">subject</span><span class="sxs-lookup"><span data-stu-id="a0ec0-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="a0ec0-160">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-160">Member</span></span> |
| [<span data-ttu-id="a0ec0-161">to</span><span class="sxs-lookup"><span data-stu-id="a0ec0-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0ec0-162">Membro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-162">Member</span></span> |
| [<span data-ttu-id="a0ec0-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a0ec0-164">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-164">Method</span></span> |
| [<span data-ttu-id="a0ec0-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a0ec0-166">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-166">Method</span></span> |
| [<span data-ttu-id="a0ec0-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a0ec0-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="a0ec0-168">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-168">Method</span></span> |
| [<span data-ttu-id="a0ec0-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a0ec0-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="a0ec0-170">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-170">Method</span></span> |
| [<span data-ttu-id="a0ec0-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="a0ec0-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="a0ec0-172">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-172">Method</span></span> |
| [<span data-ttu-id="a0ec0-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a0ec0-174">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-174">Method</span></span> |
| [<span data-ttu-id="a0ec0-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a0ec0-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a0ec0-176">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-176">Method</span></span> |
| [<span data-ttu-id="a0ec0-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a0ec0-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a0ec0-178">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-178">Method</span></span> |
| [<span data-ttu-id="a0ec0-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a0ec0-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a0ec0-180">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-180">Method</span></span> |
| [<span data-ttu-id="a0ec0-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a0ec0-182">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-182">Method</span></span> |
| [<span data-ttu-id="a0ec0-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a0ec0-184">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-184">Method</span></span> |
| [<span data-ttu-id="a0ec0-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a0ec0-186">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-186">Method</span></span> |
| [<span data-ttu-id="a0ec0-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a0ec0-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a0ec0-188">Método</span><span class="sxs-lookup"><span data-stu-id="a0ec0-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a0ec0-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-189">Example</span></span>

<span data-ttu-id="a0ec0-190">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a0ec0-191">Membros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-192">anexos: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="a0ec0-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="a0ec0-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-195">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a0ec0-196">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-197">Type</span></span>

*   <span data-ttu-id="a0ec0-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="a0ec0-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-199">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-199">Requirements</span></span>

|<span data-ttu-id="a0ec0-200">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-200">Requirement</span></span>| <span data-ttu-id="a0ec0-201">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-202">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-203">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-203">1.0</span></span>|
|[<span data-ttu-id="a0ec0-204">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-205">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-206">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-207">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-208">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-208">Example</span></span>

<span data-ttu-id="a0ec0-209">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-210">CCO: [destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-211">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a0ec0-212">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-213">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-213">Type</span></span>

*   [<span data-ttu-id="a0ec0-214">Destinatários</span><span class="sxs-lookup"><span data-stu-id="a0ec0-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-215">Requirements</span></span>

|<span data-ttu-id="a0ec0-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-216">Requirement</span></span>| <span data-ttu-id="a0ec0-217">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-219">1.1</span><span class="sxs-lookup"><span data-stu-id="a0ec0-219">1.1</span></span>|
|[<span data-ttu-id="a0ec0-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-221">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-223">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="a0ec0-225">corpo: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-226">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-227">Type</span></span>

*   [<span data-ttu-id="a0ec0-228">Body</span><span class="sxs-lookup"><span data-stu-id="a0ec0-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-229">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-229">Requirements</span></span>

|<span data-ttu-id="a0ec0-230">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-230">Requirement</span></span>| <span data-ttu-id="a0ec0-231">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-232">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-233">1.1</span><span class="sxs-lookup"><span data-stu-id="a0ec0-233">1.1</span></span>|
|[<span data-ttu-id="a0ec0-234">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-235">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-236">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-238">Example</span></span>

<span data-ttu-id="a0ec0-239">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="a0ec0-240">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-241">[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="a0ec0-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-242">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a0ec0-243">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-244">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-244">Read mode</span></span>

<span data-ttu-id="a0ec0-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-247">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-247">Compose mode</span></span>

<span data-ttu-id="a0ec0-248">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-249">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-249">Type</span></span>

*   <span data-ttu-id="a0ec0-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-251">Requirements</span></span>

|<span data-ttu-id="a0ec0-252">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-252">Requirement</span></span>| <span data-ttu-id="a0ec0-253">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-254">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-255">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-255">1.0</span></span>|
|[<span data-ttu-id="a0ec0-256">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-257">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-258">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-259">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-259">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="a0ec0-260">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="a0ec0-261">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a0ec0-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a0ec0-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-266">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-266">Type</span></span>

*   <span data-ttu-id="a0ec0-267">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-268">Requirements</span></span>

|<span data-ttu-id="a0ec0-269">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-269">Requirement</span></span>| <span data-ttu-id="a0ec0-270">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-272">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-272">1.0</span></span>|
|[<span data-ttu-id="a0ec0-273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-274">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-275">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-277">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="a0ec0-278">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="a0ec0-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="a0ec0-p110">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-281">Type</span></span>

*   <span data-ttu-id="a0ec0-282">Data</span><span class="sxs-lookup"><span data-stu-id="a0ec0-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-283">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-283">Requirements</span></span>

|<span data-ttu-id="a0ec0-284">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-284">Requirement</span></span>| <span data-ttu-id="a0ec0-285">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-286">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-287">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-287">1.0</span></span>|
|[<span data-ttu-id="a0ec0-288">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-289">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-290">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-291">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-292">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-292">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="a0ec0-293">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="a0ec0-293">dateTimeModified: Date</span></span>

<span data-ttu-id="a0ec0-294">Obtém a data e a hora em que um item foi alterado pela última vez.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="a0ec0-295">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-296">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-297">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-297">Type</span></span>

*   <span data-ttu-id="a0ec0-298">Data</span><span class="sxs-lookup"><span data-stu-id="a0ec0-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-299">Requirements</span></span>

|<span data-ttu-id="a0ec0-300">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-300">Requirement</span></span>| <span data-ttu-id="a0ec0-301">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-303">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-303">1.0</span></span>|
|[<span data-ttu-id="a0ec0-304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-305">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-307">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-308">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="a0ec0-309">fim: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-310">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a0ec0-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-313">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-313">Read mode</span></span>

<span data-ttu-id="a0ec0-314">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-314">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-315">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-315">Compose mode</span></span>

<span data-ttu-id="a0ec0-316">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a0ec0-317">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a0ec0-318">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a0ec0-319">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-319">Type</span></span>

*   <span data-ttu-id="a0ec0-320">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-321">Requirements</span></span>

|<span data-ttu-id="a0ec0-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-322">Requirement</span></span>| <span data-ttu-id="a0ec0-323">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-325">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-325">1.0</span></span>|
|[<span data-ttu-id="a0ec0-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-327">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-328">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-329">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-329">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-330">de: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a0ec0-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-335">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-336">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-336">Type</span></span>

*   [<span data-ttu-id="a0ec0-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0ec0-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-338">Requirements</span></span>

|<span data-ttu-id="a0ec0-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-339">Requirement</span></span>| <span data-ttu-id="a0ec0-340">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-342">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-342">1.0</span></span>|
|[<span data-ttu-id="a0ec0-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-344">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-345">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-346">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="a0ec0-348">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a0ec0-348">internetMessageId: String</span></span>

<span data-ttu-id="a0ec0-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-351">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-351">Type</span></span>

*   <span data-ttu-id="a0ec0-352">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-353">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-353">Requirements</span></span>

|<span data-ttu-id="a0ec0-354">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-354">Requirement</span></span>| <span data-ttu-id="a0ec0-355">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-356">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-357">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-357">1.0</span></span>|
|[<span data-ttu-id="a0ec0-358">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-359">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-360">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-361">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-362">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-362">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="a0ec0-363">doclass: String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-363">itemClass: String</span></span>

<span data-ttu-id="a0ec0-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a0ec0-p117">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a0ec0-368">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-368">Type</span></span> | <span data-ttu-id="a0ec0-369">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-369">Description</span></span> | <span data-ttu-id="a0ec0-370">classe de item</span><span class="sxs-lookup"><span data-stu-id="a0ec0-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a0ec0-371">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="a0ec0-371">Appointment items</span></span> | <span data-ttu-id="a0ec0-372">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="a0ec0-373">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-373">Message items</span></span> | <span data-ttu-id="a0ec0-374">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a0ec0-375">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-376">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-376">Type</span></span>

*   <span data-ttu-id="a0ec0-377">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-378">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-378">Requirements</span></span>

|<span data-ttu-id="a0ec0-379">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-379">Requirement</span></span>| <span data-ttu-id="a0ec0-380">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-381">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-382">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-382">1.0</span></span>|
|[<span data-ttu-id="a0ec0-383">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-384">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-385">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-386">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-387">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-387">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a0ec0-388">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-388">(nullable) itemId: String</span></span>

<span data-ttu-id="a0ec0-389">Obtém o identificador do item dos Serviços Web do Exchange para o item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="a0ec0-390">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-391">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a0ec0-392">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a0ec0-393">Antes de fazer chamadas da API REST usando esse valor, ele deve ser `Office.context.mailbox.convertToRestId`convertido usando o, que está disponível a partir do conjunto de requisitos 1,3.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="a0ec0-394">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-395">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-395">Type</span></span>

*   <span data-ttu-id="a0ec0-396">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-397">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-397">Requirements</span></span>

|<span data-ttu-id="a0ec0-398">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-398">Requirement</span></span>| <span data-ttu-id="a0ec0-399">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-400">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-401">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-401">1.0</span></span>|
|[<span data-ttu-id="a0ec0-402">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-403">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-404">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-405">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-406">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-406">Example</span></span>

<span data-ttu-id="a0ec0-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="a0ec0-409">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-410">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a0ec0-411">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-412">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-412">Type</span></span>

*   [<span data-ttu-id="a0ec0-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-414">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-414">Requirements</span></span>

|<span data-ttu-id="a0ec0-415">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-415">Requirement</span></span>| <span data-ttu-id="a0ec0-416">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-417">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-418">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-418">1.0</span></span>|
|[<span data-ttu-id="a0ec0-419">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-420">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-421">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-422">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-423">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-423">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="a0ec0-424">local: cadeia de caracteres | [Local](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-425">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-426">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-426">Read mode</span></span>

<span data-ttu-id="a0ec0-427">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-428">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-428">Compose mode</span></span>

<span data-ttu-id="a0ec0-429">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-430">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-430">Type</span></span>

*   <span data-ttu-id="a0ec0-431">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-432">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-432">Requirements</span></span>

|<span data-ttu-id="a0ec0-433">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-433">Requirement</span></span>| <span data-ttu-id="a0ec0-434">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-435">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-436">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-436">1.0</span></span>|
|[<span data-ttu-id="a0ec0-437">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-438">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-439">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-440">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-440">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a0ec0-441">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a0ec0-441">normalizedSubject: String</span></span>

<span data-ttu-id="a0ec0-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a0ec0-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-446">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-446">Type</span></span>

*   <span data-ttu-id="a0ec0-447">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-448">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-448">Requirements</span></span>

|<span data-ttu-id="a0ec0-449">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-449">Requirement</span></span>| <span data-ttu-id="a0ec0-450">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-451">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-452">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-452">1.0</span></span>|
|[<span data-ttu-id="a0ec0-453">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-454">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-455">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-456">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-457">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-457">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-458">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="a0ec0-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-459">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a0ec0-460">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-461">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-461">Read mode</span></span>

<span data-ttu-id="a0ec0-462">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-463">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-463">Compose mode</span></span>

<span data-ttu-id="a0ec0-464">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-465">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-465">Type</span></span>

*   <span data-ttu-id="a0ec0-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-467">Requirements</span></span>

|<span data-ttu-id="a0ec0-468">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-468">Requirement</span></span>| <span data-ttu-id="a0ec0-469">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-471">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-471">1.0</span></span>|
|[<span data-ttu-id="a0ec0-472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-473">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-474">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-475">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-475">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-476">organizador: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-479">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-479">Type</span></span>

*   [<span data-ttu-id="a0ec0-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0ec0-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-481">Requirements</span></span>

|<span data-ttu-id="a0ec0-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-482">Requirement</span></span>| <span data-ttu-id="a0ec0-483">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-485">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-485">1.0</span></span>|
|[<span data-ttu-id="a0ec0-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-487">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-489">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-490">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-490">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-491">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="a0ec0-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-492">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a0ec0-493">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-494">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-494">Read mode</span></span>

<span data-ttu-id="a0ec0-495">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-496">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-496">Compose mode</span></span>

<span data-ttu-id="a0ec0-497">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-498">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-498">Type</span></span>

*   <span data-ttu-id="a0ec0-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-500">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-500">Requirements</span></span>

|<span data-ttu-id="a0ec0-501">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-501">Requirement</span></span>| <span data-ttu-id="a0ec0-502">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-503">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-504">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-504">1.0</span></span>|
|[<span data-ttu-id="a0ec0-505">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-506">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-507">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-508">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-508">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-509">remetente: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a0ec0-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-514">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a0ec0-515">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-515">Type</span></span>

*   [<span data-ttu-id="a0ec0-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0ec0-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="a0ec0-517">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-517">Requirements</span></span>

|<span data-ttu-id="a0ec0-518">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-518">Requirement</span></span>| <span data-ttu-id="a0ec0-519">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-520">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-521">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-521">1.0</span></span>|
|[<span data-ttu-id="a0ec0-522">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-523">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-524">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-525">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-526">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-526">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="a0ec0-527">Início: data | [Tempo](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-528">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a0ec0-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-531">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-531">Read mode</span></span>

<span data-ttu-id="a0ec0-532">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-532">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-533">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-533">Compose mode</span></span>

<span data-ttu-id="a0ec0-534">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a0ec0-535">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="a0ec0-536">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a0ec0-537">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-537">Type</span></span>

*   <span data-ttu-id="a0ec0-538">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-539">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-539">Requirements</span></span>

|<span data-ttu-id="a0ec0-540">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-540">Requirement</span></span>| <span data-ttu-id="a0ec0-541">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-542">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-543">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-543">1.0</span></span>|
|[<span data-ttu-id="a0ec0-544">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-545">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-546">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-547">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-547">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="a0ec0-548">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-549">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a0ec0-550">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-551">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-551">Read mode</span></span>

<span data-ttu-id="a0ec0-p130">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-554">Compose mode</span></span>

<span data-ttu-id="a0ec0-555">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-556">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-556">Type</span></span>

*   <span data-ttu-id="a0ec0-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-558">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-558">Requirements</span></span>

|<span data-ttu-id="a0ec0-559">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-559">Requirement</span></span>| <span data-ttu-id="a0ec0-560">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-561">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-562">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-562">1.0</span></span>|
|[<span data-ttu-id="a0ec0-563">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-564">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-565">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-566">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-566">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="a0ec0-567">para: Array. <[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.2) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0ec0-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="a0ec0-568">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a0ec0-569">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0ec0-570">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a0ec0-570">Read mode</span></span>

<span data-ttu-id="a0ec0-p132">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0ec0-573">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a0ec0-573">Compose mode</span></span>

<span data-ttu-id="a0ec0-574">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0ec0-575">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-575">Type</span></span>

*   <span data-ttu-id="a0ec0-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-577">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-577">Requirements</span></span>

|<span data-ttu-id="a0ec0-578">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-578">Requirement</span></span>| <span data-ttu-id="a0ec0-579">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-580">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-581">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-581">1.0</span></span>|
|[<span data-ttu-id="a0ec0-582">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-583">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-584">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-585">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a0ec0-586">Métodos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a0ec0-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a0ec0-588">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a0ec0-589">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a0ec0-590">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-591">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-591">Parameters</span></span>

|<span data-ttu-id="a0ec0-592">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-592">Name</span></span>| <span data-ttu-id="a0ec0-593">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-593">Type</span></span>| <span data-ttu-id="a0ec0-594">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-594">Attributes</span></span>| <span data-ttu-id="a0ec0-595">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a0ec0-596">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-596">String</span></span>||<span data-ttu-id="a0ec0-p133">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a0ec0-599">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-599">String</span></span>||<span data-ttu-id="a0ec0-p134">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a0ec0-602">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-602">Object</span></span>| <span data-ttu-id="a0ec0-603">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-603">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-604">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a0ec0-605">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-605">Object</span></span>| <span data-ttu-id="a0ec0-606">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-606">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-607">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a0ec0-608">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-608">function</span></span>| <span data-ttu-id="a0ec0-609">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-609">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-610">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0ec0-611">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a0ec0-612">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0ec0-613">Erros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-613">Errors</span></span>

| <span data-ttu-id="a0ec0-614">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-614">Error code</span></span> | <span data-ttu-id="a0ec0-615">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a0ec0-616">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a0ec0-617">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a0ec0-618">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-619">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-619">Requirements</span></span>

|<span data-ttu-id="a0ec0-620">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-620">Requirement</span></span>| <span data-ttu-id="a0ec0-621">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-622">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-623">1.1</span><span class="sxs-lookup"><span data-stu-id="a0ec0-623">1.1</span></span>|
|[<span data-ttu-id="a0ec0-624">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0ec0-626">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-627">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-628">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a0ec0-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a0ec0-630">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a0ec0-p135">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a0ec0-634">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a0ec0-635">Se o suplemento do Office estiver em execução no Outlook na Web, o `addItemAttachmentAsync` método poderá anexar itens a itens diferentes do item que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-636">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-636">Parameters</span></span>

|<span data-ttu-id="a0ec0-637">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-637">Name</span></span>| <span data-ttu-id="a0ec0-638">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-638">Type</span></span>| <span data-ttu-id="a0ec0-639">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-639">Attributes</span></span>| <span data-ttu-id="a0ec0-640">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a0ec0-641">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-641">String</span></span>||<span data-ttu-id="a0ec0-p136">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a0ec0-644">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a0ec0-644">String</span></span>||<span data-ttu-id="a0ec0-645">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-645">The subject of the item to be attached.</span></span> <span data-ttu-id="a0ec0-646">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a0ec0-647">Object</span><span class="sxs-lookup"><span data-stu-id="a0ec0-647">Object</span></span>| <span data-ttu-id="a0ec0-648">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-648">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-649">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a0ec0-650">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-650">Object</span></span>| <span data-ttu-id="a0ec0-651">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-651">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-652">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a0ec0-653">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-653">function</span></span>| <span data-ttu-id="a0ec0-654">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-654">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-655">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0ec0-656">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a0ec0-657">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0ec0-658">Erros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-658">Errors</span></span>

| <span data-ttu-id="a0ec0-659">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-659">Error code</span></span> | <span data-ttu-id="a0ec0-660">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a0ec0-661">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-662">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-662">Requirements</span></span>

|<span data-ttu-id="a0ec0-663">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-663">Requirement</span></span>| <span data-ttu-id="a0ec0-664">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-665">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-666">1.1</span><span class="sxs-lookup"><span data-stu-id="a0ec0-666">1.1</span></span>|
|[<span data-ttu-id="a0ec0-667">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0ec0-669">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-670">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-671">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-671">Example</span></span>

<span data-ttu-id="a0ec0-672">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="a0ec0-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="a0ec0-674">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-675">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0ec0-676">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a0ec0-677">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a0ec0-678">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="a0ec0-679">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="a0ec0-680">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-681">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-681">Parameters</span></span>

|<span data-ttu-id="a0ec0-682">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-682">Name</span></span>| <span data-ttu-id="a0ec0-683">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-683">Type</span></span>| <span data-ttu-id="a0ec0-684">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a0ec0-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a0ec0-685">String &#124; Object</span></span>| |<span data-ttu-id="a0ec0-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a0ec0-688">**OU**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-688">**OR**</span></span><br/><span data-ttu-id="a0ec0-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a0ec0-691">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-691">String</span></span> | <span data-ttu-id="a0ec0-692">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-692">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a0ec0-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a0ec0-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-696">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-697">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a0ec0-698">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-698">String</span></span> | | <span data-ttu-id="a0ec0-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a0ec0-701">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a0ec0-701">String</span></span> | | <span data-ttu-id="a0ec0-702">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a0ec0-703">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-703">String</span></span> | | <span data-ttu-id="a0ec0-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a0ec0-706">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-706">String</span></span> | | <span data-ttu-id="a0ec0-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a0ec0-710">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-710">function</span></span> | <span data-ttu-id="a0ec0-711">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-711">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-712">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-713">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-713">Requirements</span></span>

|<span data-ttu-id="a0ec0-714">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-714">Requirement</span></span>| <span data-ttu-id="a0ec0-715">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-716">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-717">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-717">1.0</span></span>|
|[<span data-ttu-id="a0ec0-718">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-719">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-720">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-721">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0ec0-722">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-722">Examples</span></span>

<span data-ttu-id="a0ec0-723">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a0ec0-724">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-724">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a0ec0-725">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-725">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a0ec0-726">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a0ec0-727">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a0ec0-728">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="a0ec0-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="a0ec0-730">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-731">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0ec0-732">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de 3 colunas e um formulário pop-up no modo de exibição de 2 ou 1 colunas.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a0ec0-733">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a0ec0-734">Quando os `formData.attachments` anexos são especificados no parâmetro, o Outlook na Web e clientes da área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="a0ec0-735">Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="a0ec0-736">Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-737">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-737">Parameters</span></span>

|<span data-ttu-id="a0ec0-738">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-738">Name</span></span>| <span data-ttu-id="a0ec0-739">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-739">Type</span></span>| <span data-ttu-id="a0ec0-740">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a0ec0-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a0ec0-741">String &#124; Object</span></span>| | <span data-ttu-id="a0ec0-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a0ec0-744">**OU**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-744">**OR**</span></span><br/><span data-ttu-id="a0ec0-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a0ec0-747">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-747">String</span></span> | <span data-ttu-id="a0ec0-748">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-748">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a0ec0-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a0ec0-752">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-752">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-753">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a0ec0-754">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-754">String</span></span> | | <span data-ttu-id="a0ec0-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a0ec0-757">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a0ec0-757">String</span></span> | | <span data-ttu-id="a0ec0-758">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a0ec0-759">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-759">String</span></span> | | <span data-ttu-id="a0ec0-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a0ec0-762">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-762">String</span></span> | | <span data-ttu-id="a0ec0-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a0ec0-766">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-766">function</span></span> | <span data-ttu-id="a0ec0-767">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-767">&lt;optional&gt;</span></span> | <span data-ttu-id="a0ec0-768">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-769">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-769">Requirements</span></span>

|<span data-ttu-id="a0ec0-770">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-770">Requirement</span></span>| <span data-ttu-id="a0ec0-771">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-772">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-773">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-773">1.0</span></span>|
|[<span data-ttu-id="a0ec0-774">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-775">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-776">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-777">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0ec0-778">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-778">Examples</span></span>

<span data-ttu-id="a0ec0-779">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a0ec0-780">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-780">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a0ec0-781">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-781">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a0ec0-782">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a0ec0-783">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a0ec0-784">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="a0ec0-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="a0ec0-786">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-787">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-788">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-788">Requirements</span></span>

|<span data-ttu-id="a0ec0-789">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-789">Requirement</span></span>| <span data-ttu-id="a0ec0-790">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-791">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-792">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-792">1.0</span></span>|
|[<span data-ttu-id="a0ec0-793">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-794">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-795">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-796">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-797">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-797">Returns:</span></span>

<span data-ttu-id="a0ec0-798">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="a0ec0-799">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-799">Example</span></span>

<span data-ttu-id="a0ec0-800">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-800">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="a0ec0-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="a0ec0-802">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-803">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-804">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-804">Parameters</span></span>

|<span data-ttu-id="a0ec0-805">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-805">Name</span></span>| <span data-ttu-id="a0ec0-806">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-806">Type</span></span>| <span data-ttu-id="a0ec0-807">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a0ec0-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="a0ec0-809">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0ec0-810">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-810">Requirements</span></span>

|<span data-ttu-id="a0ec0-811">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-811">Requirement</span></span>| <span data-ttu-id="a0ec0-812">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-813">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-814">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-814">1.0</span></span>|
|[<span data-ttu-id="a0ec0-815">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-816">Restrito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-816">Restricted</span></span>|
|[<span data-ttu-id="a0ec0-817">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-818">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-819">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-819">Returns:</span></span>

<span data-ttu-id="a0ec0-820">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a0ec0-821">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a0ec0-822">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a0ec0-823">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a0ec0-824">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a0ec0-824">Value of `entityType`</span></span> | <span data-ttu-id="a0ec0-825">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="a0ec0-825">Type of objects in returned array</span></span> | <span data-ttu-id="a0ec0-826">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="a0ec0-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a0ec0-827">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-827">String</span></span> | <span data-ttu-id="a0ec0-828">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a0ec0-829">Contato</span><span class="sxs-lookup"><span data-stu-id="a0ec0-829">Contact</span></span> | <span data-ttu-id="a0ec0-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a0ec0-831">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-831">String</span></span> | <span data-ttu-id="a0ec0-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a0ec0-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a0ec0-833">MeetingSuggestion</span></span> | <span data-ttu-id="a0ec0-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a0ec0-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a0ec0-835">PhoneNumber</span></span> | <span data-ttu-id="a0ec0-836">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a0ec0-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a0ec0-837">TaskSuggestion</span></span> | <span data-ttu-id="a0ec0-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a0ec0-839">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-839">String</span></span> | <span data-ttu-id="a0ec0-840">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a0ec0-840">**Restricted**</span></span> |

<span data-ttu-id="a0ec0-841">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="a0ec0-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="a0ec0-842">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-842">Example</span></span>

<span data-ttu-id="a0ec0-843">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="a0ec0-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="a0ec0-845">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-846">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0ec0-847">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-848">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-848">Parameters</span></span>

|<span data-ttu-id="a0ec0-849">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-849">Name</span></span>| <span data-ttu-id="a0ec0-850">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-850">Type</span></span>| <span data-ttu-id="a0ec0-851">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a0ec0-852">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-852">String</span></span>|<span data-ttu-id="a0ec0-853">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0ec0-854">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-854">Requirements</span></span>

|<span data-ttu-id="a0ec0-855">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-855">Requirement</span></span>| <span data-ttu-id="a0ec0-856">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-857">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-858">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-858">1.0</span></span>|
|[<span data-ttu-id="a0ec0-859">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-860">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-861">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-862">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-863">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-863">Returns:</span></span>

<span data-ttu-id="a0ec0-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a0ec0-866">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="a0ec0-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="a0ec0-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a0ec0-868">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-869">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0ec0-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a0ec0-873">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a0ec0-874">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="a0ec0-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0ec0-877">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-877">Requirements</span></span>

|<span data-ttu-id="a0ec0-878">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-878">Requirement</span></span>| <span data-ttu-id="a0ec0-879">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-880">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-881">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-881">1.0</span></span>|
|[<span data-ttu-id="a0ec0-882">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-883">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-884">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-885">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-886">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-886">Returns:</span></span>

<span data-ttu-id="a0ec0-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="a0ec0-889">Tipo: objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-889">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="a0ec0-890">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-890">Example</span></span>

<span data-ttu-id="a0ec0-891">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="a0ec0-891">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a0ec0-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a0ec0-893">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-893">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ec0-894">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-894">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0ec0-895">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-895">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a0ec0-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-898">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-898">Parameters</span></span>

|<span data-ttu-id="a0ec0-899">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-899">Name</span></span>| <span data-ttu-id="a0ec0-900">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-900">Type</span></span>| <span data-ttu-id="a0ec0-901">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-901">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a0ec0-902">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-902">String</span></span>|<span data-ttu-id="a0ec0-903">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-903">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0ec0-904">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-904">Requirements</span></span>

|<span data-ttu-id="a0ec0-905">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-905">Requirement</span></span>| <span data-ttu-id="a0ec0-906">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-906">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-907">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-907">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-908">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-908">1.0</span></span>|
|[<span data-ttu-id="a0ec0-909">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-909">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-910">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-910">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-911">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-911">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-912">Read</span><span class="sxs-lookup"><span data-stu-id="a0ec0-912">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-913">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-913">Returns:</span></span>

<span data-ttu-id="a0ec0-914">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-914">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="a0ec0-915">Tipo: cadeia de caracteres de matriz. < ></span><span class="sxs-lookup"><span data-stu-id="a0ec0-915">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="a0ec0-916">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-916">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a0ec0-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a0ec0-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a0ec0-918">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-918">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a0ec0-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-921">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-921">Parameters</span></span>

|<span data-ttu-id="a0ec0-922">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-922">Name</span></span>| <span data-ttu-id="a0ec0-923">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-923">Type</span></span>| <span data-ttu-id="a0ec0-924">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-924">Attributes</span></span>| <span data-ttu-id="a0ec0-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-925">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="a0ec0-926">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-926">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a0ec0-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="a0ec0-930">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-930">Object</span></span>| <span data-ttu-id="a0ec0-931">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-931">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-932">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-932">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a0ec0-933">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-933">Object</span></span>| <span data-ttu-id="a0ec0-934">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-934">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-935">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-935">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a0ec0-936">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-936">function</span></span>||<span data-ttu-id="a0ec0-937">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a0ec0-938">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-938">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a0ec0-939">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-939">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0ec0-940">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-940">Requirements</span></span>

|<span data-ttu-id="a0ec0-941">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-941">Requirement</span></span>| <span data-ttu-id="a0ec0-942">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-943">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-944">1.2</span><span class="sxs-lookup"><span data-stu-id="a0ec0-944">1.2</span></span>|
|[<span data-ttu-id="a0ec0-945">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-946">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-947">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-948">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-948">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0ec0-949">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a0ec0-949">Returns:</span></span>

<span data-ttu-id="a0ec0-950">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-950">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="a0ec0-951">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-951">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a0ec0-952">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-952">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a0ec0-953">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-953">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a0ec0-954">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-954">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a0ec0-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-958">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-958">Parameters</span></span>

|<span data-ttu-id="a0ec0-959">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-959">Name</span></span>| <span data-ttu-id="a0ec0-960">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-960">Type</span></span>| <span data-ttu-id="a0ec0-961">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-961">Attributes</span></span>| <span data-ttu-id="a0ec0-962">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-962">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a0ec0-963">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-963">function</span></span>||<span data-ttu-id="a0ec0-964">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-964">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a0ec0-965">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-965">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a0ec0-966">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-966">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a0ec0-967">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-967">Object</span></span>| <span data-ttu-id="a0ec0-968">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-968">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-969">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-969">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a0ec0-970">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-970">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0ec0-971">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-971">Requirements</span></span>

|<span data-ttu-id="a0ec0-972">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-972">Requirement</span></span>| <span data-ttu-id="a0ec0-973">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-973">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-974">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-974">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-975">1.0</span><span class="sxs-lookup"><span data-stu-id="a0ec0-975">1.0</span></span>|
|[<span data-ttu-id="a0ec0-976">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-976">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-977">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-977">ReadItem</span></span>|
|[<span data-ttu-id="a0ec0-978">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a0ec0-978">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-979">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a0ec0-979">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-980">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-980">Example</span></span>

<span data-ttu-id="a0ec0-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a0ec0-984">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0ec0-984">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a0ec0-985">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-985">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a0ec0-986">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-986">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a0ec0-987">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-987">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a0ec0-988">No Outlook na Web e dispositivos móveis, o identificador de anexo é válido somente dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-988">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a0ec0-989">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-989">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-990">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-990">Parameters</span></span>

|<span data-ttu-id="a0ec0-991">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-991">Name</span></span>| <span data-ttu-id="a0ec0-992">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-992">Type</span></span>| <span data-ttu-id="a0ec0-993">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-993">Attributes</span></span>| <span data-ttu-id="a0ec0-994">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-994">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a0ec0-995">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-995">String</span></span>||<span data-ttu-id="a0ec0-996">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-996">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="a0ec0-997">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-997">Object</span></span>| <span data-ttu-id="a0ec0-998">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-998">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-999">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-999">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a0ec0-1000">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1000">Object</span></span>| <span data-ttu-id="a0ec0-1001">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-1002">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1002">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a0ec0-1003">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1003">function</span></span>| <span data-ttu-id="a0ec0-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-1005">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1005">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0ec0-1006">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1006">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0ec0-1007">Erros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1007">Errors</span></span>

| <span data-ttu-id="a0ec0-1008">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1008">Error code</span></span> | <span data-ttu-id="a0ec0-1009">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1009">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a0ec0-1010">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1010">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-1011">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1011">Requirements</span></span>

|<span data-ttu-id="a0ec0-1012">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1012">Requirement</span></span>| <span data-ttu-id="a0ec0-1013">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-1014">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-1015">1.1</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1015">1.1</span></span>|
|[<span data-ttu-id="a0ec0-1016">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0ec0-1018">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-1019">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1019">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-1020">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1020">Example</span></span>

<span data-ttu-id="a0ec0-1021">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1021">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a0ec0-1022">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1022">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a0ec0-1023">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1023">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a0ec0-p166">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0ec0-1027">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1027">Parameters</span></span>

|<span data-ttu-id="a0ec0-1028">Nome</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1028">Name</span></span>| <span data-ttu-id="a0ec0-1029">Tipo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1029">Type</span></span>| <span data-ttu-id="a0ec0-1030">Atributos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1030">Attributes</span></span>| <span data-ttu-id="a0ec0-1031">Descrição</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1031">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a0ec0-1032">String</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1032">String</span></span>||<span data-ttu-id="a0ec0-p167">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="a0ec0-1036">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1036">Object</span></span>| <span data-ttu-id="a0ec0-1037">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-1038">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1038">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a0ec0-1039">Objeto</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1039">Object</span></span>| <span data-ttu-id="a0ec0-1040">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-1041">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1041">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a0ec0-1042">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1042">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a0ec0-1043">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="a0ec0-1044">Se `text`, o estilo atual é aplicado no Outlook na Web e clientes da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1044">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="a0ec0-1045">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1045">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a0ec0-1046">Se `html` e o campo oferecer suporte a HTML (o assunto não), o estilo atual será aplicado no Outlook na Web e o estilo padrão será aplicado nos clientes da área de trabalho do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1046">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="a0ec0-1047">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1047">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a0ec0-1048">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1048">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="a0ec0-1049">function</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1049">function</span></span>||<span data-ttu-id="a0ec0-1050">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1050">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a0ec0-1051">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1051">Requirements</span></span>

|<span data-ttu-id="a0ec0-1052">Requisito</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1052">Requirement</span></span>| <span data-ttu-id="a0ec0-1053">Valor</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0ec0-1054">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0ec0-1055">1.2</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1055">1.2</span></span>|
|[<span data-ttu-id="a0ec0-1056">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0ec0-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0ec0-1058">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0ec0-1059">Escrever</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0ec0-1060">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a0ec0-1060">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
