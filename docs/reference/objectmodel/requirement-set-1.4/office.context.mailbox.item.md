---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,4
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 0644a7f6c6d9c6532ad4126653a30c53867635ad
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066260"
---
# <a name="item"></a><span data-ttu-id="945e8-102">item</span><span class="sxs-lookup"><span data-stu-id="945e8-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="945e8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="945e8-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="945e8-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="945e8-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-106">Requirements</span></span>

|<span data-ttu-id="945e8-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-107">Requirement</span></span>| <span data-ttu-id="945e8-108">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-110">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-110">1.0</span></span>|
|[<span data-ttu-id="945e8-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="945e8-112">Restricted</span></span>|
|[<span data-ttu-id="945e8-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="945e8-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="945e8-115">Members and methods</span></span>

| <span data-ttu-id="945e8-116">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-116">Member</span></span> | <span data-ttu-id="945e8-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="945e8-118">attachments</span><span class="sxs-lookup"><span data-stu-id="945e8-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="945e8-119">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-119">Member</span></span> |
| [<span data-ttu-id="945e8-120">bcc</span><span class="sxs-lookup"><span data-stu-id="945e8-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="945e8-121">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-121">Member</span></span> |
| [<span data-ttu-id="945e8-122">body</span><span class="sxs-lookup"><span data-stu-id="945e8-122">body</span></span>](#body-body) | <span data-ttu-id="945e8-123">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-123">Member</span></span> |
| [<span data-ttu-id="945e8-124">cc</span><span class="sxs-lookup"><span data-stu-id="945e8-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="945e8-125">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-125">Member</span></span> |
| [<span data-ttu-id="945e8-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="945e8-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="945e8-127">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-127">Member</span></span> |
| [<span data-ttu-id="945e8-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="945e8-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="945e8-129">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-129">Member</span></span> |
| [<span data-ttu-id="945e8-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="945e8-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="945e8-131">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-131">Member</span></span> |
| [<span data-ttu-id="945e8-132">end</span><span class="sxs-lookup"><span data-stu-id="945e8-132">end</span></span>](#end-datetime) | <span data-ttu-id="945e8-133">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-133">Member</span></span> |
| [<span data-ttu-id="945e8-134">from</span><span class="sxs-lookup"><span data-stu-id="945e8-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="945e8-135">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-135">Member</span></span> |
| [<span data-ttu-id="945e8-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="945e8-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="945e8-137">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-137">Member</span></span> |
| [<span data-ttu-id="945e8-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="945e8-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="945e8-139">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-139">Member</span></span> |
| [<span data-ttu-id="945e8-140">itemId</span><span class="sxs-lookup"><span data-stu-id="945e8-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="945e8-141">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-141">Member</span></span> |
| [<span data-ttu-id="945e8-142">itemType</span><span class="sxs-lookup"><span data-stu-id="945e8-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="945e8-143">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-143">Member</span></span> |
| [<span data-ttu-id="945e8-144">location</span><span class="sxs-lookup"><span data-stu-id="945e8-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="945e8-145">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-145">Member</span></span> |
| [<span data-ttu-id="945e8-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="945e8-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="945e8-147">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-147">Member</span></span> |
| [<span data-ttu-id="945e8-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="945e8-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="945e8-149">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-149">Member</span></span> |
| [<span data-ttu-id="945e8-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="945e8-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="945e8-151">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-151">Member</span></span> |
| [<span data-ttu-id="945e8-152">organizer</span><span class="sxs-lookup"><span data-stu-id="945e8-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="945e8-153">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-153">Member</span></span> |
| [<span data-ttu-id="945e8-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="945e8-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="945e8-155">Member</span><span class="sxs-lookup"><span data-stu-id="945e8-155">Member</span></span> |
| [<span data-ttu-id="945e8-156">sender</span><span class="sxs-lookup"><span data-stu-id="945e8-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="945e8-157">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-157">Member</span></span> |
| [<span data-ttu-id="945e8-158">start</span><span class="sxs-lookup"><span data-stu-id="945e8-158">start</span></span>](#start-datetime) | <span data-ttu-id="945e8-159">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-159">Member</span></span> |
| [<span data-ttu-id="945e8-160">subject</span><span class="sxs-lookup"><span data-stu-id="945e8-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="945e8-161">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-161">Member</span></span> |
| [<span data-ttu-id="945e8-162">to</span><span class="sxs-lookup"><span data-stu-id="945e8-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="945e8-163">Membro</span><span class="sxs-lookup"><span data-stu-id="945e8-163">Member</span></span> |
| [<span data-ttu-id="945e8-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="945e8-165">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-165">Method</span></span> |
| [<span data-ttu-id="945e8-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="945e8-167">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-167">Method</span></span> |
| [<span data-ttu-id="945e8-168">close</span><span class="sxs-lookup"><span data-stu-id="945e8-168">close</span></span>](#close) | <span data-ttu-id="945e8-169">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-169">Method</span></span> |
| [<span data-ttu-id="945e8-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="945e8-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="945e8-171">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-171">Method</span></span> |
| [<span data-ttu-id="945e8-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="945e8-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="945e8-173">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-173">Method</span></span> |
| [<span data-ttu-id="945e8-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="945e8-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="945e8-175">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-175">Method</span></span> |
| [<span data-ttu-id="945e8-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="945e8-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="945e8-177">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-177">Method</span></span> |
| [<span data-ttu-id="945e8-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="945e8-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="945e8-179">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-179">Method</span></span> |
| [<span data-ttu-id="945e8-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="945e8-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="945e8-181">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-181">Method</span></span> |
| [<span data-ttu-id="945e8-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="945e8-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="945e8-183">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-183">Method</span></span> |
| [<span data-ttu-id="945e8-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="945e8-185">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-185">Method</span></span> |
| [<span data-ttu-id="945e8-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="945e8-187">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-187">Method</span></span> |
| [<span data-ttu-id="945e8-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="945e8-189">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-189">Method</span></span> |
| [<span data-ttu-id="945e8-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="945e8-191">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-191">Method</span></span> |
| [<span data-ttu-id="945e8-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="945e8-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="945e8-193">Método</span><span class="sxs-lookup"><span data-stu-id="945e8-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="945e8-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-194">Example</span></span>

<span data-ttu-id="945e8-195">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="945e8-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="945e8-196">Members</span><span class="sxs-lookup"><span data-stu-id="945e8-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="945e8-197">anexos: Matriz.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="945e8-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="945e8-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-200">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="945e8-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="945e8-201">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="945e8-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-202">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-202">Type</span></span>

*   <span data-ttu-id="945e8-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="945e8-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-204">Requirements</span></span>

|<span data-ttu-id="945e8-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-205">Requirement</span></span>| <span data-ttu-id="945e8-206">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-208">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-208">1.0</span></span>|
|[<span data-ttu-id="945e8-209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-210">ReadItem</span></span>|
|[<span data-ttu-id="945e8-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-212">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-213">Example</span></span>

<span data-ttu-id="945e8-214">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="945e8-215">cco :[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-216">Obtém um objeto que fornece métodos para obter ou atualizar a linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="945e8-217">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="945e8-217">Compose mode only.</span></span>

<span data-ttu-id="945e8-218">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-219">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="945e8-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="945e8-220">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="945e8-221">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="945e8-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-222">Type</span></span>

*   [<span data-ttu-id="945e8-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="945e8-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-224">Requirements</span></span>

|<span data-ttu-id="945e8-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-225">Requirement</span></span>| <span data-ttu-id="945e8-226">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-228">1.1</span><span class="sxs-lookup"><span data-stu-id="945e8-228">1.1</span></span>|
|[<span data-ttu-id="945e8-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-230">ReadItem</span></span>|
|[<span data-ttu-id="945e8-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="945e8-234">corpo: [Corpo](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="945e8-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-236">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-236">Type</span></span>

*   [<span data-ttu-id="945e8-237">Body</span><span class="sxs-lookup"><span data-stu-id="945e8-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-238">Requirements</span></span>

|<span data-ttu-id="945e8-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-239">Requirement</span></span>| <span data-ttu-id="945e8-240">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-242">1.1</span><span class="sxs-lookup"><span data-stu-id="945e8-242">1.1</span></span>|
|[<span data-ttu-id="945e8-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-244">ReadItem</span></span>|
|[<span data-ttu-id="945e8-245">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-247">Example</span></span>

<span data-ttu-id="945e8-248">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="945e8-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="945e8-249">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="945e8-250">cc : Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-251">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="945e8-252">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-253">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-253">Read mode</span></span>

<span data-ttu-id="945e8-254">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="945e8-255">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-256">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-257">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-257">Compose mode</span></span>

<span data-ttu-id="945e8-258">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="945e8-259">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-260">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="945e8-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="945e8-261">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="945e8-262">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="945e8-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="945e8-263">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-263">Type</span></span>

*   <span data-ttu-id="945e8-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-265">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-265">Requirements</span></span>

|<span data-ttu-id="945e8-266">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-266">Requirement</span></span>| <span data-ttu-id="945e8-267">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-268">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-269">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-269">1.0</span></span>|
|[<span data-ttu-id="945e8-270">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-271">ReadItem</span></span>|
|[<span data-ttu-id="945e8-272">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-273">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="945e8-274">(anulável) conversationId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="945e8-275">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="945e8-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="945e8-p109">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="945e8-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="945e8-p110">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="945e8-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-280">Type</span></span>

*   <span data-ttu-id="945e8-281">String</span><span class="sxs-lookup"><span data-stu-id="945e8-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-282">Requirements</span></span>

|<span data-ttu-id="945e8-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-283">Requirement</span></span>| <span data-ttu-id="945e8-284">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-286">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-286">1.0</span></span>|
|[<span data-ttu-id="945e8-287">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-288">ReadItem</span></span>|
|[<span data-ttu-id="945e8-289">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-290">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-291">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="945e8-292">dateTimeCreated: Data</span><span class="sxs-lookup"><span data-stu-id="945e8-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="945e8-p111">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-295">Type</span></span>

*   <span data-ttu-id="945e8-296">Data</span><span class="sxs-lookup"><span data-stu-id="945e8-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-297">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-297">Requirements</span></span>

|<span data-ttu-id="945e8-298">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-298">Requirement</span></span>| <span data-ttu-id="945e8-299">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-300">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-301">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-301">1.0</span></span>|
|[<span data-ttu-id="945e8-302">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-303">ReadItem</span></span>|
|[<span data-ttu-id="945e8-304">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-305">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-306">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="945e8-307">dateTimeModified: Data</span><span class="sxs-lookup"><span data-stu-id="945e8-307">dateTimeModified: Date</span></span>

<span data-ttu-id="945e8-p112">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-310">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-311">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-311">Type</span></span>

*   <span data-ttu-id="945e8-312">Data</span><span class="sxs-lookup"><span data-stu-id="945e8-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-313">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-313">Requirements</span></span>

|<span data-ttu-id="945e8-314">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-314">Requirement</span></span>| <span data-ttu-id="945e8-315">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-316">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-317">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-317">1.0</span></span>|
|[<span data-ttu-id="945e8-318">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-319">ReadItem</span></span>|
|[<span data-ttu-id="945e8-320">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-321">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-322">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="945e8-323">fim: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-324">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="945e8-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="945e8-p113">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="945e8-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-327">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-327">Read mode</span></span>

<span data-ttu-id="945e8-328">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="945e8-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-329">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-329">Compose mode</span></span>

<span data-ttu-id="945e8-330">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="945e8-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="945e8-331">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="945e8-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="945e8-332">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="945e8-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="945e8-333">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-333">Type</span></span>

*   <span data-ttu-id="945e8-334">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-335">Requirements</span></span>

|<span data-ttu-id="945e8-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-336">Requirement</span></span>| <span data-ttu-id="945e8-337">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-339">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-339">1.0</span></span>|
|[<span data-ttu-id="945e8-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-341">ReadItem</span></span>|
|[<span data-ttu-id="945e8-342">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-343">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="945e8-344">De:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-p114">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="945e8-p115">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="945e8-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-349">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="945e8-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-350">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-350">Type</span></span>

*   [<span data-ttu-id="945e8-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="945e8-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-352">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-352">Requirements</span></span>

|<span data-ttu-id="945e8-353">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-353">Requirement</span></span>| <span data-ttu-id="945e8-354">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-355">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-356">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-356">1.0</span></span>|
|[<span data-ttu-id="945e8-357">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-358">ReadItem</span></span>|
|[<span data-ttu-id="945e8-359">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-360">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-361">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="945e8-362">internetMessageId: Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-362">internetMessageId: String</span></span>

<span data-ttu-id="945e8-p116">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-365">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-365">Type</span></span>

*   <span data-ttu-id="945e8-366">String</span><span class="sxs-lookup"><span data-stu-id="945e8-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-367">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-367">Requirements</span></span>

|<span data-ttu-id="945e8-368">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-368">Requirement</span></span>| <span data-ttu-id="945e8-369">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-370">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-371">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-371">1.0</span></span>|
|[<span data-ttu-id="945e8-372">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-373">ReadItem</span></span>|
|[<span data-ttu-id="945e8-374">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-375">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-376">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="945e8-377">itemClass: Cadeia de caracteres </span><span class="sxs-lookup"><span data-stu-id="945e8-377">itemClass: String</span></span>

<span data-ttu-id="945e8-p117">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="945e8-p118">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="945e8-382">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-382">Type</span></span> | <span data-ttu-id="945e8-383">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-383">Description</span></span> | <span data-ttu-id="945e8-384">classe de item</span><span class="sxs-lookup"><span data-stu-id="945e8-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="945e8-385">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="945e8-385">Appointment items</span></span> | <span data-ttu-id="945e8-386">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="945e8-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="945e8-387">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="945e8-387">Message items</span></span> | <span data-ttu-id="945e8-388">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="945e8-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="945e8-389">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="945e8-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-390">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-390">Type</span></span>

*   <span data-ttu-id="945e8-391">String</span><span class="sxs-lookup"><span data-stu-id="945e8-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-392">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-392">Requirements</span></span>

|<span data-ttu-id="945e8-393">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-393">Requirement</span></span>| <span data-ttu-id="945e8-394">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-395">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-396">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-396">1.0</span></span>|
|[<span data-ttu-id="945e8-397">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-398">ReadItem</span></span>|
|[<span data-ttu-id="945e8-399">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-400">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-401">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="945e8-402">(anulável) itemId: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-402">(nullable) itemId: String</span></span>

<span data-ttu-id="945e8-p119">Obtém o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-405">O identificador retornado pela propriedade `itemId` é o mesmo que o [identificador do item dos Serviços Web do Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="945e8-405">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="945e8-406">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="945e8-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="945e8-407">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="945e8-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="945e8-408">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="945e8-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="945e8-p121">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-411">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-411">Type</span></span>

*   <span data-ttu-id="945e8-412">String</span><span class="sxs-lookup"><span data-stu-id="945e8-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-413">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-413">Requirements</span></span>

|<span data-ttu-id="945e8-414">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-414">Requirement</span></span>| <span data-ttu-id="945e8-415">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-416">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-417">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-417">1.0</span></span>|
|[<span data-ttu-id="945e8-418">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-419">ReadItem</span></span>|
|[<span data-ttu-id="945e8-420">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-421">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-422">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-422">Example</span></span>

<span data-ttu-id="945e8-p122">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="945e8-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="945e8-425">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-426">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="945e8-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="945e8-427">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-428">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-428">Type</span></span>

*   [<span data-ttu-id="945e8-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="945e8-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-430">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-430">Requirements</span></span>

|<span data-ttu-id="945e8-431">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-431">Requirement</span></span>| <span data-ttu-id="945e8-432">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-433">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-434">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-434">1.0</span></span>|
|[<span data-ttu-id="945e8-435">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-436">ReadItem</span></span>|
|[<span data-ttu-id="945e8-437">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-438">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-439">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="945e8-440">Local: Cadeia de caracteres[Local](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-441">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-442">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-442">Read mode</span></span>

<span data-ttu-id="945e8-443">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-444">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-444">Compose mode</span></span>

<span data-ttu-id="945e8-445">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="945e8-446">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-446">Type</span></span>

*   <span data-ttu-id="945e8-447">Cadeia de caracteres | [Localização](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-448">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-448">Requirements</span></span>

|<span data-ttu-id="945e8-449">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-449">Requirement</span></span>| <span data-ttu-id="945e8-450">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-451">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-452">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-452">1.0</span></span>|
|[<span data-ttu-id="945e8-453">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-454">ReadItem</span></span>|
|[<span data-ttu-id="945e8-455">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-456">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="945e8-457">normalizedSubject: Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-457">normalizedSubject: String</span></span>

<span data-ttu-id="945e8-p123">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="945e8-p124">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="945e8-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-462">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-462">Type</span></span>

*   <span data-ttu-id="945e8-463">String</span><span class="sxs-lookup"><span data-stu-id="945e8-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-464">Requirements</span></span>

|<span data-ttu-id="945e8-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-465">Requirement</span></span>| <span data-ttu-id="945e8-466">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-468">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-468">1.0</span></span>|
|[<span data-ttu-id="945e8-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-470">ReadItem</span></span>|
|[<span data-ttu-id="945e8-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-472">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="945e8-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-475">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="945e8-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-476">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-476">Type</span></span>

*   [<span data-ttu-id="945e8-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="945e8-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-478">Requirements</span></span>

|<span data-ttu-id="945e8-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-479">Requirement</span></span>| <span data-ttu-id="945e8-480">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-482">1.3</span><span class="sxs-lookup"><span data-stu-id="945e8-482">1.3</span></span>|
|[<span data-ttu-id="945e8-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-484">ReadItem</span></span>|
|[<span data-ttu-id="945e8-485">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-486">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-487">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="945e8-488">optionalAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-489">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="945e8-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="945e8-490">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-491">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-491">Read mode</span></span>

<span data-ttu-id="945e8-492">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="945e8-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="945e8-493">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-494">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-495">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-495">Compose mode</span></span>

<span data-ttu-id="945e8-496">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="945e8-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="945e8-497">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-498">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="945e8-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="945e8-499">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="945e8-500">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="945e8-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="945e8-501">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-501">Type</span></span>

*   <span data-ttu-id="945e8-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-503">Requirements</span></span>

|<span data-ttu-id="945e8-504">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-504">Requirement</span></span>| <span data-ttu-id="945e8-505">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-506">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-507">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-507">1.0</span></span>|
|[<span data-ttu-id="945e8-508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-509">ReadItem</span></span>|
|[<span data-ttu-id="945e8-510">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-511">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="945e8-512">organizador:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-p128">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-515">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-515">Type</span></span>

*   [<span data-ttu-id="945e8-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="945e8-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-517">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-517">Requirements</span></span>

|<span data-ttu-id="945e8-518">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-518">Requirement</span></span>| <span data-ttu-id="945e8-519">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-520">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-521">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-521">1.0</span></span>|
|[<span data-ttu-id="945e8-522">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-523">ReadItem</span></span>|
|[<span data-ttu-id="945e8-524">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-525">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-526">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="945e8-527">requiredAttendees: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-528">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="945e8-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="945e8-529">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-530">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-530">Read mode</span></span>

<span data-ttu-id="945e8-531">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="945e8-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="945e8-532">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-533">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-534">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-534">Compose mode</span></span>

<span data-ttu-id="945e8-535">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="945e8-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="945e8-536">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-537">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="945e8-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="945e8-538">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="945e8-539">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="945e8-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="945e8-540">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-540">Type</span></span>

*   <span data-ttu-id="945e8-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-542">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-542">Requirements</span></span>

|<span data-ttu-id="945e8-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-543">Requirement</span></span>| <span data-ttu-id="945e8-544">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-546">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-546">1.0</span></span>|
|[<span data-ttu-id="945e8-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-548">ReadItem</span></span>|
|[<span data-ttu-id="945e8-549">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-550">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="945e8-551">remetente :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-p132">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="945e8-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="945e8-p133">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="945e8-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-556">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="945e8-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="945e8-557">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-557">Type</span></span>

*   [<span data-ttu-id="945e8-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="945e8-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="945e8-559">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-559">Requirements</span></span>

|<span data-ttu-id="945e8-560">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-560">Requirement</span></span>| <span data-ttu-id="945e8-561">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-562">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-563">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-563">1.0</span></span>|
|[<span data-ttu-id="945e8-564">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-565">ReadItem</span></span>|
|[<span data-ttu-id="945e8-566">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-567">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-568">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="945e8-569">início: Data|[Hora](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-570">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="945e8-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="945e8-p134">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="945e8-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-573">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-573">Read mode</span></span>

<span data-ttu-id="945e8-574">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="945e8-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-575">Compose mode</span></span>

<span data-ttu-id="945e8-576">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="945e8-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="945e8-577">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="945e8-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="945e8-578">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="945e8-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="945e8-579">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-579">Type</span></span>

*   <span data-ttu-id="945e8-580">Data | [Hora](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-581">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-581">Requirements</span></span>

|<span data-ttu-id="945e8-582">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-582">Requirement</span></span>| <span data-ttu-id="945e8-583">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-584">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-585">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-585">1.0</span></span>|
|[<span data-ttu-id="945e8-586">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-587">ReadItem</span></span>|
|[<span data-ttu-id="945e8-588">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-589">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="945e8-590">Assunto: Cadeia de caracteres|[Assunto](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-591">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="945e8-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="945e8-592">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="945e8-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-593">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-593">Read mode</span></span>

<span data-ttu-id="945e8-p135">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="945e8-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-596">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-596">Compose mode</span></span>

<span data-ttu-id="945e8-597">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="945e8-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="945e8-598">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-598">Type</span></span>

*   <span data-ttu-id="945e8-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-600">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-600">Requirements</span></span>

|<span data-ttu-id="945e8-601">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-601">Requirement</span></span>| <span data-ttu-id="945e8-602">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-603">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-604">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-604">1.0</span></span>|
|[<span data-ttu-id="945e8-605">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-606">ReadItem</span></span>|
|[<span data-ttu-id="945e8-607">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-608">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="945e8-609">para: Matriz.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Destinatários](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="945e8-610">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="945e8-611">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="945e8-612">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="945e8-612">Read mode</span></span>

<span data-ttu-id="945e8-613">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="945e8-614">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-615">No entanto, no Windows e Mac, você pode ter o máximo de 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="945e8-616">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="945e8-616">Compose mode</span></span>

<span data-ttu-id="945e8-617">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="945e8-618">Por padrão, o conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="945e8-619">No entanto, no Windows e no Mac, os seguintes limites se aplicam.</span><span class="sxs-lookup"><span data-stu-id="945e8-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="945e8-620">Tenha no máximo 500 membros.</span><span class="sxs-lookup"><span data-stu-id="945e8-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="945e8-621">Defina o máximo de 100 membros por chamada, até 500 no total.</span><span class="sxs-lookup"><span data-stu-id="945e8-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="945e8-622">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-622">Type</span></span>

*   <span data-ttu-id="945e8-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-624">Requirements</span></span>

|<span data-ttu-id="945e8-625">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-625">Requirement</span></span>| <span data-ttu-id="945e8-626">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-627">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-628">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-628">1.0</span></span>|
|[<span data-ttu-id="945e8-629">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-630">ReadItem</span></span>|
|[<span data-ttu-id="945e8-631">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-632">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="945e8-633">Métodos</span><span class="sxs-lookup"><span data-stu-id="945e8-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="945e8-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="945e8-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="945e8-635">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="945e8-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="945e8-636">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="945e8-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="945e8-637">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="945e8-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-638">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-638">Parameters</span></span>

|<span data-ttu-id="945e8-639">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-639">Name</span></span>| <span data-ttu-id="945e8-640">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-640">Type</span></span>| <span data-ttu-id="945e8-641">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-641">Attributes</span></span>| <span data-ttu-id="945e8-642">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="945e8-643">String</span><span class="sxs-lookup"><span data-stu-id="945e8-643">String</span></span>||<span data-ttu-id="945e8-p139">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="945e8-646">String</span><span class="sxs-lookup"><span data-stu-id="945e8-646">String</span></span>||<span data-ttu-id="945e8-p140">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="945e8-649">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-649">Object</span></span>| <span data-ttu-id="945e8-650">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-650">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-651">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-651">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-652">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-652">Object</span></span>| <span data-ttu-id="945e8-653">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-653">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-654">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-654">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="945e8-655">function</span><span class="sxs-lookup"><span data-stu-id="945e8-655">function</span></span>| <span data-ttu-id="945e8-656">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-656">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-657">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-657">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="945e8-658">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="945e8-658">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="945e8-659">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="945e8-659">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="945e8-660">Erros</span><span class="sxs-lookup"><span data-stu-id="945e8-660">Errors</span></span>

| <span data-ttu-id="945e8-661">Código de erro</span><span class="sxs-lookup"><span data-stu-id="945e8-661">Error code</span></span> | <span data-ttu-id="945e8-662">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-662">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="945e8-663">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="945e8-663">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="945e8-664">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="945e8-664">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="945e8-665">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="945e8-665">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-666">Requirements</span></span>

|<span data-ttu-id="945e8-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-667">Requirement</span></span>| <span data-ttu-id="945e8-668">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-670">1.1</span><span class="sxs-lookup"><span data-stu-id="945e8-670">1.1</span></span>|
|[<span data-ttu-id="945e8-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-672">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="945e8-672">ReadWriteItem</span></span>|
|[<span data-ttu-id="945e8-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-674">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-674">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-675">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-675">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="945e8-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="945e8-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="945e8-677">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-677">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="945e8-p141">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="945e8-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="945e8-681">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="945e8-681">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="945e8-682">Se o Suplemento do Office estiver em execução no Outlook na Web, o método `addItemAttachmentAsync` pode anexar itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="945e8-682">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-683">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-683">Parameters</span></span>

|<span data-ttu-id="945e8-684">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-684">Name</span></span>| <span data-ttu-id="945e8-685">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-685">Type</span></span>| <span data-ttu-id="945e8-686">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-686">Attributes</span></span>| <span data-ttu-id="945e8-687">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-687">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="945e8-688">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-688">String</span></span>||<span data-ttu-id="945e8-p142">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="945e8-691">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-691">String</span></span>||<span data-ttu-id="945e8-692">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="945e8-692">The subject of the item to be attached.</span></span> <span data-ttu-id="945e8-693">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-693">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="945e8-694">Object</span><span class="sxs-lookup"><span data-stu-id="945e8-694">Object</span></span>| <span data-ttu-id="945e8-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-695">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-696">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-696">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-697">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-697">Object</span></span>| <span data-ttu-id="945e8-698">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-698">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-699">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-699">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="945e8-700">function</span><span class="sxs-lookup"><span data-stu-id="945e8-700">function</span></span>| <span data-ttu-id="945e8-701">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-701">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-702">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="945e8-703">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="945e8-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="945e8-704">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="945e8-704">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="945e8-705">Erros</span><span class="sxs-lookup"><span data-stu-id="945e8-705">Errors</span></span>

| <span data-ttu-id="945e8-706">Código de erro</span><span class="sxs-lookup"><span data-stu-id="945e8-706">Error code</span></span> | <span data-ttu-id="945e8-707">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-707">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="945e8-708">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="945e8-708">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-709">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-709">Requirements</span></span>

|<span data-ttu-id="945e8-710">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-710">Requirement</span></span>| <span data-ttu-id="945e8-711">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-712">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-713">1.1</span><span class="sxs-lookup"><span data-stu-id="945e8-713">1.1</span></span>|
|[<span data-ttu-id="945e8-714">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-714">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-715">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="945e8-715">ReadWriteItem</span></span>|
|[<span data-ttu-id="945e8-716">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-716">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-717">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-717">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-718">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-718">Example</span></span>

<span data-ttu-id="945e8-719">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="945e8-719">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="945e8-720">close()</span><span class="sxs-lookup"><span data-stu-id="945e8-720">close()</span></span>

<span data-ttu-id="945e8-721">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="945e8-721">Closes the current item that is being composed.</span></span>

<span data-ttu-id="945e8-p144">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="945e8-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-724">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="945e8-724">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="945e8-725">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="945e8-725">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-726">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-726">Requirements</span></span>

|<span data-ttu-id="945e8-727">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-727">Requirement</span></span>| <span data-ttu-id="945e8-728">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-729">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-730">1.3</span><span class="sxs-lookup"><span data-stu-id="945e8-730">1.3</span></span>|
|[<span data-ttu-id="945e8-731">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-732">Restrito</span><span class="sxs-lookup"><span data-stu-id="945e8-732">Restricted</span></span>|
|[<span data-ttu-id="945e8-733">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-734">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-734">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="945e8-735">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="945e8-735">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="945e8-736">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="945e8-736">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-737">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="945e8-738">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="945e8-738">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="945e8-739">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="945e8-739">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="945e8-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="945e8-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-743">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-743">Parameters</span></span>

|<span data-ttu-id="945e8-744">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-744">Name</span></span>| <span data-ttu-id="945e8-745">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-745">Type</span></span>| <span data-ttu-id="945e8-746">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-746">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="945e8-747">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="945e8-747">String &#124; Object</span></span>| |<span data-ttu-id="945e8-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="945e8-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="945e8-750">**OU**</span><span class="sxs-lookup"><span data-stu-id="945e8-750">**OR**</span></span><br/><span data-ttu-id="945e8-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="945e8-753">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-753">String</span></span> | <span data-ttu-id="945e8-754">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-754">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="945e8-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="945e8-757">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-757">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="945e8-758">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-758">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-759">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="945e8-759">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="945e8-760">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-760">String</span></span> | | <span data-ttu-id="945e8-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="945e8-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="945e8-763">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-763">String</span></span> | | <span data-ttu-id="945e8-764">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="945e8-764">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="945e8-765">String</span><span class="sxs-lookup"><span data-stu-id="945e8-765">String</span></span> | | <span data-ttu-id="945e8-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="945e8-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="945e8-768">String</span><span class="sxs-lookup"><span data-stu-id="945e8-768">String</span></span> | | <span data-ttu-id="945e8-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="945e8-772">function</span><span class="sxs-lookup"><span data-stu-id="945e8-772">function</span></span> | <span data-ttu-id="945e8-773">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-773">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-774">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-774">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-775">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-775">Requirements</span></span>

|<span data-ttu-id="945e8-776">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-776">Requirement</span></span>| <span data-ttu-id="945e8-777">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-777">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-778">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-778">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-779">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-779">1.0</span></span>|
|[<span data-ttu-id="945e8-780">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-780">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-781">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-781">ReadItem</span></span>|
|[<span data-ttu-id="945e8-782">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-782">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-783">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-783">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="945e8-784">Exemplos</span><span class="sxs-lookup"><span data-stu-id="945e8-784">Examples</span></span>

<span data-ttu-id="945e8-785">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="945e8-785">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="945e8-786">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="945e8-786">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="945e8-787">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="945e8-787">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="945e8-788">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="945e8-788">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="945e8-789">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="945e8-789">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="945e8-790">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-790">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="945e8-791">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="945e8-791">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="945e8-792">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="945e8-792">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-793">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-793">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="945e8-794">No Outlook na Web, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="945e8-794">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="945e8-795">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="945e8-795">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="945e8-p152">Quando os anexos são especificados no parâmetro `formData.attachments`, os clientes do Outlook na Web e do Outlook para área de trabalho tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="945e8-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-799">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-799">Parameters</span></span>

|<span data-ttu-id="945e8-800">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-800">Name</span></span>| <span data-ttu-id="945e8-801">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-801">Type</span></span>| <span data-ttu-id="945e8-802">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-802">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="945e8-803">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="945e8-803">String &#124; Object</span></span>| | <span data-ttu-id="945e8-p153">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="945e8-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="945e8-806">**OU**</span><span class="sxs-lookup"><span data-stu-id="945e8-806">**OR**</span></span><br/><span data-ttu-id="945e8-p154">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="945e8-809">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-809">String</span></span> | <span data-ttu-id="945e8-810">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-810">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-p155">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="945e8-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="945e8-813">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-813">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="945e8-814">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-814">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-815">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="945e8-815">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="945e8-816">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-816">String</span></span> | | <span data-ttu-id="945e8-p156">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="945e8-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="945e8-819">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-819">String</span></span> | | <span data-ttu-id="945e8-820">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="945e8-820">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="945e8-821">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-821">String</span></span> | | <span data-ttu-id="945e8-p157">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="945e8-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="945e8-824">String</span><span class="sxs-lookup"><span data-stu-id="945e8-824">String</span></span> | | <span data-ttu-id="945e8-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="945e8-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="945e8-828">function</span><span class="sxs-lookup"><span data-stu-id="945e8-828">function</span></span> | <span data-ttu-id="945e8-829">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-829">&lt;optional&gt;</span></span> | <span data-ttu-id="945e8-830">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-830">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-831">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-831">Requirements</span></span>

|<span data-ttu-id="945e8-832">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-832">Requirement</span></span>| <span data-ttu-id="945e8-833">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-834">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-835">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-835">1.0</span></span>|
|[<span data-ttu-id="945e8-836">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-837">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-837">ReadItem</span></span>|
|[<span data-ttu-id="945e8-838">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-839">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-839">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="945e8-840">Exemplos</span><span class="sxs-lookup"><span data-stu-id="945e8-840">Examples</span></span>

<span data-ttu-id="945e8-841">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="945e8-841">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="945e8-842">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="945e8-842">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="945e8-843">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="945e8-843">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="945e8-844">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="945e8-844">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="945e8-845">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="945e8-845">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="945e8-846">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-846">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="945e8-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="945e8-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="945e8-848">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="945e8-848">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-849">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-849">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-850">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-850">Requirements</span></span>

|<span data-ttu-id="945e8-851">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-851">Requirement</span></span>| <span data-ttu-id="945e8-852">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-853">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-854">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-854">1.0</span></span>|
|[<span data-ttu-id="945e8-855">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-855">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-856">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-856">ReadItem</span></span>|
|[<span data-ttu-id="945e8-857">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-857">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-858">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-858">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-859">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-859">Returns:</span></span>

<span data-ttu-id="945e8-860">Tipo: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="945e8-860">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="945e8-861">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-861">Example</span></span>

<span data-ttu-id="945e8-862">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-862">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="945e8-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="945e8-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="945e8-864">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="945e8-864">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-865">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-866">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-866">Parameters</span></span>

|<span data-ttu-id="945e8-867">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-867">Name</span></span>| <span data-ttu-id="945e8-868">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-868">Type</span></span>| <span data-ttu-id="945e8-869">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-869">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="945e8-870">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="945e8-870">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="945e8-871">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="945e8-871">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-872">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-872">Requirements</span></span>

|<span data-ttu-id="945e8-873">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-873">Requirement</span></span>| <span data-ttu-id="945e8-874">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-874">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-875">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-875">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-876">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-876">1.0</span></span>|
|[<span data-ttu-id="945e8-877">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-877">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-878">Restrito</span><span class="sxs-lookup"><span data-stu-id="945e8-878">Restricted</span></span>|
|[<span data-ttu-id="945e8-879">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-879">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-880">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-880">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-881">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-881">Returns:</span></span>

<span data-ttu-id="945e8-882">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="945e8-882">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="945e8-883">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="945e8-883">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="945e8-884">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="945e8-884">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="945e8-885">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-885">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="945e8-886">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="945e8-886">Value of `entityType`</span></span> | <span data-ttu-id="945e8-887">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="945e8-887">Type of objects in returned array</span></span> | <span data-ttu-id="945e8-888">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="945e8-888">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="945e8-889">String</span><span class="sxs-lookup"><span data-stu-id="945e8-889">String</span></span> | <span data-ttu-id="945e8-890">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="945e8-890">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="945e8-891">Contato</span><span class="sxs-lookup"><span data-stu-id="945e8-891">Contact</span></span> | <span data-ttu-id="945e8-892">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="945e8-892">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="945e8-893">String</span><span class="sxs-lookup"><span data-stu-id="945e8-893">String</span></span> | <span data-ttu-id="945e8-894">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="945e8-894">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="945e8-895">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="945e8-895">MeetingSuggestion</span></span> | <span data-ttu-id="945e8-896">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="945e8-896">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="945e8-897">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="945e8-897">PhoneNumber</span></span> | <span data-ttu-id="945e8-898">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="945e8-898">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="945e8-899">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="945e8-899">TaskSuggestion</span></span> | <span data-ttu-id="945e8-900">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="945e8-900">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="945e8-901">String</span><span class="sxs-lookup"><span data-stu-id="945e8-901">String</span></span> | <span data-ttu-id="945e8-902">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="945e8-902">**Restricted**</span></span> |

<span data-ttu-id="945e8-903">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="945e8-903">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="945e8-904">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-904">Example</span></span>

<span data-ttu-id="945e8-905">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="945e8-905">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="945e8-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="945e8-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="945e8-907">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="945e8-907">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-908">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-908">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="945e8-909">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="945e8-909">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-910">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-910">Parameters</span></span>

|<span data-ttu-id="945e8-911">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-911">Name</span></span>| <span data-ttu-id="945e8-912">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-912">Type</span></span>| <span data-ttu-id="945e8-913">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-913">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="945e8-914">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-914">String</span></span>|<span data-ttu-id="945e8-915">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="945e8-915">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-916">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-916">Requirements</span></span>

|<span data-ttu-id="945e8-917">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-917">Requirement</span></span>| <span data-ttu-id="945e8-918">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-918">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-919">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-919">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-920">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-920">1.0</span></span>|
|[<span data-ttu-id="945e8-921">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-921">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-922">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-922">ReadItem</span></span>|
|[<span data-ttu-id="945e8-923">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-923">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-924">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-924">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-925">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-925">Returns:</span></span>

<span data-ttu-id="945e8-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="945e8-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="945e8-928">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="945e8-928">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="945e8-929">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="945e8-929">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="945e8-930">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="945e8-930">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-931">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="945e8-p161">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="945e8-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="945e8-935">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="945e8-935">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="945e8-936">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="945e8-936">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="945e8-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="945e8-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="945e8-940">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-940">Requirements</span></span>

|<span data-ttu-id="945e8-941">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-941">Requirement</span></span>| <span data-ttu-id="945e8-942">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-943">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-944">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-944">1.0</span></span>|
|[<span data-ttu-id="945e8-945">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-946">ReadItem</span></span>|
|[<span data-ttu-id="945e8-947">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-948">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-948">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-949">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-949">Returns:</span></span>

<span data-ttu-id="945e8-p163">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="945e8-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="945e8-952">Tipo: Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-952">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="945e8-953">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-953">Example</span></span>

<span data-ttu-id="945e8-954">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="945e8-954">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="945e8-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="945e8-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="945e8-956">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="945e8-956">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-957">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="945e8-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="945e8-958">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="945e8-958">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="945e8-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="945e8-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-961">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-961">Parameters</span></span>

|<span data-ttu-id="945e8-962">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-962">Name</span></span>| <span data-ttu-id="945e8-963">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-963">Type</span></span>| <span data-ttu-id="945e8-964">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-964">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="945e8-965">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="945e8-965">String</span></span>|<span data-ttu-id="945e8-966">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="945e8-966">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-967">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-967">Requirements</span></span>

|<span data-ttu-id="945e8-968">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-968">Requirement</span></span>| <span data-ttu-id="945e8-969">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-969">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-970">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-970">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-971">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-971">1.0</span></span>|
|[<span data-ttu-id="945e8-972">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-972">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-973">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-973">ReadItem</span></span>|
|[<span data-ttu-id="945e8-974">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-974">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-975">Read</span><span class="sxs-lookup"><span data-stu-id="945e8-975">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-976">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-976">Returns:</span></span>

<span data-ttu-id="945e8-977">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="945e8-977">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="945e8-978">Tipo: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="945e8-978">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="945e8-979">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-979">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="945e8-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="945e8-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="945e8-981">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-981">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="945e8-982">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará uma cadeia de caracteres vazia para os dados selecionados.</span><span class="sxs-lookup"><span data-stu-id="945e8-982">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="945e8-983">Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="945e8-983">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-984">No Outlook na Web, o método retorna a cadeia de caracteres “null” se nenhum texto for selecionado, mas o cursor estiver no corpo.</span><span class="sxs-lookup"><span data-stu-id="945e8-984">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="945e8-985">Para verificar essa situação, confira o exemplo mais adiante nesta seção.</span><span class="sxs-lookup"><span data-stu-id="945e8-985">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-986">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-986">Parameters</span></span>

|<span data-ttu-id="945e8-987">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-987">Name</span></span>| <span data-ttu-id="945e8-988">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-988">Type</span></span>| <span data-ttu-id="945e8-989">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-989">Attributes</span></span>| <span data-ttu-id="945e8-990">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-990">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="945e8-991">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="945e8-991">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="945e8-p167">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="945e8-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="945e8-995">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-995">Object</span></span>| <span data-ttu-id="945e8-996">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-996">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-997">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-997">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-998">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-998">Object</span></span>| <span data-ttu-id="945e8-999">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-999">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1000">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1000">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="945e8-1001">function</span><span class="sxs-lookup"><span data-stu-id="945e8-1001">function</span></span>||<span data-ttu-id="945e8-1002">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-1002">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="945e8-1003">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1003">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="945e8-1004">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1004">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-1005">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-1005">Requirements</span></span>

|<span data-ttu-id="945e8-1006">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-1006">Requirement</span></span>| <span data-ttu-id="945e8-1007">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-1007">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-1008">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-1008">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-1009">1.2</span><span class="sxs-lookup"><span data-stu-id="945e8-1009">1.2</span></span>|
|[<span data-ttu-id="945e8-1010">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-1010">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-1011">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-1011">ReadItem</span></span>|
|[<span data-ttu-id="945e8-1012">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-1012">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-1013">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-1013">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="945e8-1014">Retorna:</span><span class="sxs-lookup"><span data-stu-id="945e8-1014">Returns:</span></span>

<span data-ttu-id="945e8-1015">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1015">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="945e8-1016">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="945e8-1016">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="945e8-1017">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-1017">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="945e8-1018">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="945e8-1018">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="945e8-1019">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="945e8-1019">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="945e8-p169">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="945e8-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-1023">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-1023">Parameters</span></span>

|<span data-ttu-id="945e8-1024">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-1024">Name</span></span>| <span data-ttu-id="945e8-1025">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-1025">Type</span></span>| <span data-ttu-id="945e8-1026">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-1026">Attributes</span></span>| <span data-ttu-id="945e8-1027">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-1027">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="945e8-1028">function</span><span class="sxs-lookup"><span data-stu-id="945e8-1028">function</span></span>||<span data-ttu-id="945e8-1029">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="945e8-1030">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1030">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="945e8-1031">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="945e8-1031">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="945e8-1032">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1032">Object</span></span>| <span data-ttu-id="945e8-1033">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1034">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1034">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="945e8-1035">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1035">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-1036">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-1036">Requirements</span></span>

|<span data-ttu-id="945e8-1037">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-1037">Requirement</span></span>| <span data-ttu-id="945e8-1038">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-1038">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-1039">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-1039">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-1040">1.0</span><span class="sxs-lookup"><span data-stu-id="945e8-1040">1.0</span></span>|
|[<span data-ttu-id="945e8-1041">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-1041">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-1042">ReadItem</span><span class="sxs-lookup"><span data-stu-id="945e8-1042">ReadItem</span></span>|
|[<span data-ttu-id="945e8-1043">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="945e8-1043">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-1044">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="945e8-1044">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-1045">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-1045">Example</span></span>

<span data-ttu-id="945e8-p172">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="945e8-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="945e8-1049">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="945e8-1049">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="945e8-1050">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="945e8-1050">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="945e8-1051">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="945e8-1051">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="945e8-1052">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="945e8-1052">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="945e8-1053">No Outlook na Web e em dispositivos móveis, a identificador do anexo é válido apenas durante a mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="945e8-1053">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="945e8-1054">Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1054">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-1055">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-1055">Parameters</span></span>

|<span data-ttu-id="945e8-1056">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-1056">Name</span></span>| <span data-ttu-id="945e8-1057">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-1057">Type</span></span>| <span data-ttu-id="945e8-1058">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-1058">Attributes</span></span>| <span data-ttu-id="945e8-1059">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-1059">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="945e8-1060">String</span><span class="sxs-lookup"><span data-stu-id="945e8-1060">String</span></span>||<span data-ttu-id="945e8-1061">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="945e8-1061">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="945e8-1062">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1062">Object</span></span>| <span data-ttu-id="945e8-1063">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1064">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-1065">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1065">Object</span></span>| <span data-ttu-id="945e8-1066">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1067">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="945e8-1068">function</span><span class="sxs-lookup"><span data-stu-id="945e8-1068">function</span></span>| <span data-ttu-id="945e8-1069">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1070">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="945e8-1071">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="945e8-1071">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="945e8-1072">Erros</span><span class="sxs-lookup"><span data-stu-id="945e8-1072">Errors</span></span>

| <span data-ttu-id="945e8-1073">Código de erro</span><span class="sxs-lookup"><span data-stu-id="945e8-1073">Error code</span></span> | <span data-ttu-id="945e8-1074">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-1074">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="945e8-1075">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="945e8-1075">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-1076">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-1076">Requirements</span></span>

|<span data-ttu-id="945e8-1077">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-1077">Requirement</span></span>| <span data-ttu-id="945e8-1078">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-1078">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-1079">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-1079">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-1080">1.1</span><span class="sxs-lookup"><span data-stu-id="945e8-1080">1.1</span></span>|
|[<span data-ttu-id="945e8-1081">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-1081">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-1082">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="945e8-1082">ReadWriteItem</span></span>|
|[<span data-ttu-id="945e8-1083">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-1083">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-1084">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-1084">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-1085">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-1085">Example</span></span>

<span data-ttu-id="945e8-1086">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="945e8-1086">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="945e8-1087">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="945e8-1087">saveAsync([options], callback)</span></span>

<span data-ttu-id="945e8-1088">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="945e8-1088">Asynchronously saves an item.</span></span>

<span data-ttu-id="945e8-1089">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1089">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="945e8-1090">No Outlook na Web ou no Outlook no modo online, o item é salvo no servidor.</span><span class="sxs-lookup"><span data-stu-id="945e8-1090">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="945e8-1091">No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="945e8-1091">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-1092">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="945e8-1092">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="945e8-1093">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="945e8-1093">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="945e8-p176">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="945e8-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="945e8-1097">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="945e8-1097">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="945e8-1098">O Outlook no Mac não dá suporte ao salvamento de reuniões.</span><span class="sxs-lookup"><span data-stu-id="945e8-1098">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="945e8-1099">O método `saveAsync` falha quando chamado a partir de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="945e8-1099">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="945e8-1100">Consulte [Não é possível salvar uma reunião como um rascunho no Outlook para Mac usando a API do Office JS](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="945e8-1100">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="945e8-1101">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="945e8-1101">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-1102">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-1102">Parameters</span></span>

|<span data-ttu-id="945e8-1103">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-1103">Name</span></span>| <span data-ttu-id="945e8-1104">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-1104">Type</span></span>| <span data-ttu-id="945e8-1105">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-1105">Attributes</span></span>| <span data-ttu-id="945e8-1106">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-1106">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="945e8-1107">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1107">Object</span></span>| <span data-ttu-id="945e8-1108">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1108">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1109">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-1109">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-1110">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1110">Object</span></span>| <span data-ttu-id="945e8-1111">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1111">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1112">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1112">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="945e8-1113">function</span><span class="sxs-lookup"><span data-stu-id="945e8-1113">function</span></span>||<span data-ttu-id="945e8-1114">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-1114">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="945e8-1115">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1115">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="945e8-1116">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-1116">Requirements</span></span>

|<span data-ttu-id="945e8-1117">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-1117">Requirement</span></span>| <span data-ttu-id="945e8-1118">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-1119">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-1120">1.3</span><span class="sxs-lookup"><span data-stu-id="945e8-1120">1.3</span></span>|
|[<span data-ttu-id="945e8-1121">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-1122">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="945e8-1122">ReadWriteItem</span></span>|
|[<span data-ttu-id="945e8-1123">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-1124">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-1124">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="945e8-1125">Exemplos</span><span class="sxs-lookup"><span data-stu-id="945e8-1125">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="945e8-p178">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="945e8-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="945e8-1128">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="945e8-1128">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="945e8-1129">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="945e8-1129">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="945e8-p179">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="945e8-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="945e8-1133">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="945e8-1133">Parameters</span></span>

|<span data-ttu-id="945e8-1134">Nome</span><span class="sxs-lookup"><span data-stu-id="945e8-1134">Name</span></span>| <span data-ttu-id="945e8-1135">Tipo</span><span class="sxs-lookup"><span data-stu-id="945e8-1135">Type</span></span>| <span data-ttu-id="945e8-1136">Atributos</span><span class="sxs-lookup"><span data-stu-id="945e8-1136">Attributes</span></span>| <span data-ttu-id="945e8-1137">Descrição</span><span class="sxs-lookup"><span data-stu-id="945e8-1137">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="945e8-1138">String</span><span class="sxs-lookup"><span data-stu-id="945e8-1138">String</span></span>||<span data-ttu-id="945e8-p180">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="945e8-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="945e8-1142">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1142">Object</span></span>| <span data-ttu-id="945e8-1143">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1143">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1144">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="945e8-1144">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="945e8-1145">Objeto</span><span class="sxs-lookup"><span data-stu-id="945e8-1145">Object</span></span>| <span data-ttu-id="945e8-1146">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1146">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1147">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="945e8-1147">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="945e8-1148">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="945e8-1148">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="945e8-1149">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="945e8-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="945e8-1150">Se `text`, o estilo atual é aplicado nos clientes do Outlook na Web e do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="945e8-1150">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="945e8-1151">Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="945e8-1151">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="945e8-1152">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook na Web e o estilo padrão é aplicado nos clientes do Outlook para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="945e8-1152">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="945e8-1153">Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="945e8-1153">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="945e8-1154">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="945e8-1154">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="945e8-1155">function</span><span class="sxs-lookup"><span data-stu-id="945e8-1155">function</span></span>||<span data-ttu-id="945e8-1156">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="945e8-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="945e8-1157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="945e8-1157">Requirements</span></span>

|<span data-ttu-id="945e8-1158">Requisito</span><span class="sxs-lookup"><span data-stu-id="945e8-1158">Requirement</span></span>| <span data-ttu-id="945e8-1159">Valor</span><span class="sxs-lookup"><span data-stu-id="945e8-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="945e8-1160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="945e8-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="945e8-1161">1.2</span><span class="sxs-lookup"><span data-stu-id="945e8-1161">1.2</span></span>|
|[<span data-ttu-id="945e8-1162">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="945e8-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="945e8-1163">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="945e8-1163">ReadWriteItem</span></span>|
|[<span data-ttu-id="945e8-1164">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="945e8-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="945e8-1165">Escrever</span><span class="sxs-lookup"><span data-stu-id="945e8-1165">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="945e8-1166">Exemplo</span><span class="sxs-lookup"><span data-stu-id="945e8-1166">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
