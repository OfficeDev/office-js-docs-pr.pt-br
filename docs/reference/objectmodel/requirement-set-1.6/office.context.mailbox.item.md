---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,6
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: 578e25b4fd7caf08087f24febdfd5b1877ed57bf
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706327"
---
# <a name="item"></a><span data-ttu-id="809eb-102">item</span><span class="sxs-lookup"><span data-stu-id="809eb-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="809eb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="809eb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="809eb-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="809eb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-106">Requirements</span></span>

|<span data-ttu-id="809eb-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-107">Requirement</span></span>| <span data-ttu-id="809eb-108">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-110">1.0</span></span>|
|[<span data-ttu-id="809eb-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="809eb-112">Restricted</span></span>|
|[<span data-ttu-id="809eb-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="809eb-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="809eb-115">Members and methods</span></span>

| <span data-ttu-id="809eb-116">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-116">Member</span></span> | <span data-ttu-id="809eb-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="809eb-118">attachments</span><span class="sxs-lookup"><span data-stu-id="809eb-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="809eb-119">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-119">Member</span></span> |
| [<span data-ttu-id="809eb-120">bcc</span><span class="sxs-lookup"><span data-stu-id="809eb-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="809eb-121">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-121">Member</span></span> |
| [<span data-ttu-id="809eb-122">body</span><span class="sxs-lookup"><span data-stu-id="809eb-122">body</span></span>](#body-body) | <span data-ttu-id="809eb-123">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-123">Member</span></span> |
| [<span data-ttu-id="809eb-124">cc</span><span class="sxs-lookup"><span data-stu-id="809eb-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="809eb-125">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-125">Member</span></span> |
| [<span data-ttu-id="809eb-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="809eb-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="809eb-127">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-127">Member</span></span> |
| [<span data-ttu-id="809eb-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="809eb-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="809eb-129">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-129">Member</span></span> |
| [<span data-ttu-id="809eb-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="809eb-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="809eb-131">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-131">Member</span></span> |
| [<span data-ttu-id="809eb-132">end</span><span class="sxs-lookup"><span data-stu-id="809eb-132">end</span></span>](#end-datetime) | <span data-ttu-id="809eb-133">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-133">Member</span></span> |
| [<span data-ttu-id="809eb-134">from</span><span class="sxs-lookup"><span data-stu-id="809eb-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="809eb-135">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-135">Member</span></span> |
| [<span data-ttu-id="809eb-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="809eb-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="809eb-137">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-137">Member</span></span> |
| [<span data-ttu-id="809eb-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="809eb-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="809eb-139">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-139">Member</span></span> |
| [<span data-ttu-id="809eb-140">itemId</span><span class="sxs-lookup"><span data-stu-id="809eb-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="809eb-141">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-141">Member</span></span> |
| [<span data-ttu-id="809eb-142">itemType</span><span class="sxs-lookup"><span data-stu-id="809eb-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="809eb-143">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-143">Member</span></span> |
| [<span data-ttu-id="809eb-144">location</span><span class="sxs-lookup"><span data-stu-id="809eb-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="809eb-145">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-145">Member</span></span> |
| [<span data-ttu-id="809eb-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="809eb-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="809eb-147">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-147">Member</span></span> |
| [<span data-ttu-id="809eb-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="809eb-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="809eb-149">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-149">Member</span></span> |
| [<span data-ttu-id="809eb-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="809eb-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="809eb-151">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-151">Member</span></span> |
| [<span data-ttu-id="809eb-152">organizer</span><span class="sxs-lookup"><span data-stu-id="809eb-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="809eb-153">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-153">Member</span></span> |
| [<span data-ttu-id="809eb-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="809eb-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="809eb-155">Member</span><span class="sxs-lookup"><span data-stu-id="809eb-155">Member</span></span> |
| [<span data-ttu-id="809eb-156">sender</span><span class="sxs-lookup"><span data-stu-id="809eb-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="809eb-157">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-157">Member</span></span> |
| [<span data-ttu-id="809eb-158">start</span><span class="sxs-lookup"><span data-stu-id="809eb-158">start</span></span>](#start-datetime) | <span data-ttu-id="809eb-159">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-159">Member</span></span> |
| [<span data-ttu-id="809eb-160">subject</span><span class="sxs-lookup"><span data-stu-id="809eb-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="809eb-161">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-161">Member</span></span> |
| [<span data-ttu-id="809eb-162">to</span><span class="sxs-lookup"><span data-stu-id="809eb-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="809eb-163">Membro</span><span class="sxs-lookup"><span data-stu-id="809eb-163">Member</span></span> |
| [<span data-ttu-id="809eb-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="809eb-165">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-165">Method</span></span> |
| [<span data-ttu-id="809eb-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="809eb-167">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-167">Method</span></span> |
| [<span data-ttu-id="809eb-168">close</span><span class="sxs-lookup"><span data-stu-id="809eb-168">close</span></span>](#close) | <span data-ttu-id="809eb-169">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-169">Method</span></span> |
| [<span data-ttu-id="809eb-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="809eb-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="809eb-171">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-171">Method</span></span> |
| [<span data-ttu-id="809eb-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="809eb-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="809eb-173">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-173">Method</span></span> |
| [<span data-ttu-id="809eb-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="809eb-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="809eb-175">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-175">Method</span></span> |
| [<span data-ttu-id="809eb-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="809eb-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="809eb-177">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-177">Method</span></span> |
| [<span data-ttu-id="809eb-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="809eb-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="809eb-179">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-179">Method</span></span> |
| [<span data-ttu-id="809eb-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="809eb-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="809eb-181">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-181">Method</span></span> |
| [<span data-ttu-id="809eb-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="809eb-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="809eb-183">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-183">Method</span></span> |
| [<span data-ttu-id="809eb-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="809eb-185">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-185">Method</span></span> |
| [<span data-ttu-id="809eb-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="809eb-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="809eb-187">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-187">Method</span></span> |
| [<span data-ttu-id="809eb-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="809eb-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="809eb-189">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-189">Method</span></span> |
| [<span data-ttu-id="809eb-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="809eb-191">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-191">Method</span></span> |
| [<span data-ttu-id="809eb-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="809eb-193">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-193">Method</span></span> |
| [<span data-ttu-id="809eb-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="809eb-195">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-195">Method</span></span> |
| [<span data-ttu-id="809eb-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="809eb-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="809eb-197">Método</span><span class="sxs-lookup"><span data-stu-id="809eb-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="809eb-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-198">Example</span></span>

<span data-ttu-id="809eb-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="809eb-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="809eb-200">Membros</span><span class="sxs-lookup"><span data-stu-id="809eb-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="809eb-201">anexos: Array. <[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="809eb-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="809eb-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="809eb-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="809eb-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="809eb-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-206">Type</span></span>

*   <span data-ttu-id="809eb-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="809eb-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-208">Requirements</span></span>

|<span data-ttu-id="809eb-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-209">Requirement</span></span>| <span data-ttu-id="809eb-210">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-212">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-212">1.0</span></span>|
|[<span data-ttu-id="809eb-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-214">ReadItem</span></span>|
|[<span data-ttu-id="809eb-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-216">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-217">Example</span></span>

<span data-ttu-id="809eb-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="809eb-219">CCO: [destinatários](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="809eb-219">bcc: [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="809eb-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="809eb-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="809eb-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-222">Type</span></span>

*   [<span data-ttu-id="809eb-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="809eb-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="809eb-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-224">Requirements</span></span>

|<span data-ttu-id="809eb-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-225">Requirement</span></span>| <span data-ttu-id="809eb-226">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-228">1.1</span><span class="sxs-lookup"><span data-stu-id="809eb-228">1.1</span></span>|
|[<span data-ttu-id="809eb-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-230">ReadItem</span></span>|
|[<span data-ttu-id="809eb-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="809eb-234">corpo: [Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="809eb-234">body: [Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="809eb-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="809eb-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-236">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-236">Type</span></span>

*   [<span data-ttu-id="809eb-237">Body</span><span class="sxs-lookup"><span data-stu-id="809eb-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="809eb-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-238">Requirements</span></span>

|<span data-ttu-id="809eb-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-239">Requirement</span></span>| <span data-ttu-id="809eb-240">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-242">1.1</span><span class="sxs-lookup"><span data-stu-id="809eb-242">1.1</span></span>|
|[<span data-ttu-id="809eb-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-244">ReadItem</span></span>|
|[<span data-ttu-id="809eb-245">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-247">Example</span></span>

<span data-ttu-id="809eb-248">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="809eb-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="809eb-249">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="809eb-250">[destinatários](/javascript/api/outlook_1_6/office.recipients) [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|CC: Array. <</span><span class="sxs-lookup"><span data-stu-id="809eb-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="809eb-251">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="809eb-252">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-253">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-253">Read mode</span></span>

<span data-ttu-id="809eb-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="809eb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-256">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-256">Compose mode</span></span>

<span data-ttu-id="809eb-257">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="809eb-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-258">Type</span></span>

*   <span data-ttu-id="809eb-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="809eb-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-260">Requirements</span></span>

|<span data-ttu-id="809eb-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-261">Requirement</span></span>| <span data-ttu-id="809eb-262">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-264">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-264">1.0</span></span>|
|[<span data-ttu-id="809eb-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-266">ReadItem</span></span>|
|[<span data-ttu-id="809eb-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-268">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="809eb-269">(Nullable) Conversation: String</span><span class="sxs-lookup"><span data-stu-id="809eb-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="809eb-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="809eb-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="809eb-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="809eb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="809eb-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="809eb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-275">Type</span></span>

*   <span data-ttu-id="809eb-276">String</span><span class="sxs-lookup"><span data-stu-id="809eb-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-277">Requirements</span></span>

|<span data-ttu-id="809eb-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-278">Requirement</span></span>| <span data-ttu-id="809eb-279">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-281">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-281">1.0</span></span>|
|[<span data-ttu-id="809eb-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-283">ReadItem</span></span>|
|[<span data-ttu-id="809eb-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="809eb-287">dateTimeCreated: data</span><span class="sxs-lookup"><span data-stu-id="809eb-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="809eb-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-290">Type</span></span>

*   <span data-ttu-id="809eb-291">Data</span><span class="sxs-lookup"><span data-stu-id="809eb-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-292">Requirements</span></span>

|<span data-ttu-id="809eb-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-293">Requirement</span></span>| <span data-ttu-id="809eb-294">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-296">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-296">1.0</span></span>|
|[<span data-ttu-id="809eb-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-298">ReadItem</span></span>|
|[<span data-ttu-id="809eb-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-300">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="809eb-302">dateTimeModified: data</span><span class="sxs-lookup"><span data-stu-id="809eb-302">dateTimeModified: Date</span></span>

<span data-ttu-id="809eb-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-305">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-306">Type</span></span>

*   <span data-ttu-id="809eb-307">Data</span><span class="sxs-lookup"><span data-stu-id="809eb-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-308">Requirements</span></span>

|<span data-ttu-id="809eb-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-309">Requirement</span></span>| <span data-ttu-id="809eb-310">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-312">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-312">1.0</span></span>|
|[<span data-ttu-id="809eb-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-314">ReadItem</span></span>|
|[<span data-ttu-id="809eb-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-316">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="809eb-318">fim: data | [Tempo](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="809eb-318">end: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="809eb-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="809eb-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="809eb-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="809eb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-322">Read mode</span></span>

<span data-ttu-id="809eb-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="809eb-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-324">Compose mode</span></span>

<span data-ttu-id="809eb-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="809eb-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="809eb-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="809eb-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="809eb-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="809eb-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="809eb-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-328">Type</span></span>

*   <span data-ttu-id="809eb-329">Data | [Hora](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="809eb-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-330">Requirements</span></span>

|<span data-ttu-id="809eb-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-331">Requirement</span></span>| <span data-ttu-id="809eb-332">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-334">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-334">1.0</span></span>|
|[<span data-ttu-id="809eb-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-336">ReadItem</span></span>|
|[<span data-ttu-id="809eb-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="809eb-339">de: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="809eb-339">from: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="809eb-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="809eb-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="809eb-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="809eb-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-345">Type</span></span>

*   [<span data-ttu-id="809eb-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="809eb-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="809eb-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="809eb-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-348">Requirements</span></span>

|<span data-ttu-id="809eb-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-349">Requirement</span></span>| <span data-ttu-id="809eb-350">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-352">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-352">1.0</span></span>|
|[<span data-ttu-id="809eb-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-354">ReadItem</span></span>|
|[<span data-ttu-id="809eb-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-356">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="809eb-357">internetMessageId: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="809eb-357">internetMessageId: String</span></span>

<span data-ttu-id="809eb-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-360">Type</span></span>

*   <span data-ttu-id="809eb-361">String</span><span class="sxs-lookup"><span data-stu-id="809eb-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-362">Requirements</span></span>

|<span data-ttu-id="809eb-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-363">Requirement</span></span>| <span data-ttu-id="809eb-364">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-366">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-366">1.0</span></span>|
|[<span data-ttu-id="809eb-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-368">ReadItem</span></span>|
|[<span data-ttu-id="809eb-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-370">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="809eb-372">doclass: String</span><span class="sxs-lookup"><span data-stu-id="809eb-372">itemClass: String</span></span>

<span data-ttu-id="809eb-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="809eb-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="809eb-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-377">Type</span></span> | <span data-ttu-id="809eb-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-378">Description</span></span> | <span data-ttu-id="809eb-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="809eb-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="809eb-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="809eb-380">Appointment items</span></span> | <span data-ttu-id="809eb-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="809eb-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="809eb-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="809eb-382">Message items</span></span> | <span data-ttu-id="809eb-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="809eb-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="809eb-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="809eb-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-385">Type</span></span>

*   <span data-ttu-id="809eb-386">String</span><span class="sxs-lookup"><span data-stu-id="809eb-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-387">Requirements</span></span>

|<span data-ttu-id="809eb-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-388">Requirement</span></span>| <span data-ttu-id="809eb-389">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-391">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-391">1.0</span></span>|
|[<span data-ttu-id="809eb-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-393">ReadItem</span></span>|
|[<span data-ttu-id="809eb-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-395">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="809eb-397">(Nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="809eb-397">(nullable) itemId: String</span></span>

<span data-ttu-id="809eb-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-400">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="809eb-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="809eb-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="809eb-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="809eb-402">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="809eb-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="809eb-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="809eb-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="809eb-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-406">Type</span></span>

*   <span data-ttu-id="809eb-407">String</span><span class="sxs-lookup"><span data-stu-id="809eb-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-408">Requirements</span></span>

|<span data-ttu-id="809eb-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-409">Requirement</span></span>| <span data-ttu-id="809eb-410">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-412">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-412">1.0</span></span>|
|[<span data-ttu-id="809eb-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-414">ReadItem</span></span>|
|[<span data-ttu-id="809eb-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-416">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-417">Example</span></span>

<span data-ttu-id="809eb-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="809eb-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="809eb-420">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="809eb-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="809eb-421">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="809eb-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="809eb-422">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-423">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-423">Type</span></span>

*   [<span data-ttu-id="809eb-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="809eb-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="809eb-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-425">Requirements</span></span>

|<span data-ttu-id="809eb-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-426">Requirement</span></span>| <span data-ttu-id="809eb-427">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-429">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-429">1.0</span></span>|
|[<span data-ttu-id="809eb-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-431">ReadItem</span></span>|
|[<span data-ttu-id="809eb-432">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-433">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="809eb-435">local: cadeia de caracteres | [Local](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="809eb-435">location: String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="809eb-436">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-437">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-437">Read mode</span></span>

<span data-ttu-id="809eb-438">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-439">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-439">Compose mode</span></span>

<span data-ttu-id="809eb-440">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="809eb-441">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-441">Type</span></span>

*   <span data-ttu-id="809eb-442">Cadeia de caracteres | [Localização](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="809eb-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-443">Requirements</span></span>

|<span data-ttu-id="809eb-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-444">Requirement</span></span>| <span data-ttu-id="809eb-445">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-447">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-447">1.0</span></span>|
|[<span data-ttu-id="809eb-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-449">ReadItem</span></span>|
|[<span data-ttu-id="809eb-450">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-451">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="809eb-452">normalizedSubject: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="809eb-452">normalizedSubject: String</span></span>

<span data-ttu-id="809eb-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="809eb-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="809eb-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-457">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-457">Type</span></span>

*   <span data-ttu-id="809eb-458">String</span><span class="sxs-lookup"><span data-stu-id="809eb-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-459">Requirements</span></span>

|<span data-ttu-id="809eb-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-460">Requirement</span></span>| <span data-ttu-id="809eb-461">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-463">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-463">1.0</span></span>|
|[<span data-ttu-id="809eb-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-465">ReadItem</span></span>|
|[<span data-ttu-id="809eb-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-467">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="809eb-469">notificationMessages: [notificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="809eb-469">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="809eb-470">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="809eb-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-471">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-471">Type</span></span>

*   [<span data-ttu-id="809eb-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="809eb-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="809eb-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-473">Requirements</span></span>

|<span data-ttu-id="809eb-474">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-474">Requirement</span></span>| <span data-ttu-id="809eb-475">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-477">1.3</span><span class="sxs-lookup"><span data-stu-id="809eb-477">1.3</span></span>|
|[<span data-ttu-id="809eb-478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-479">ReadItem</span></span>|
|[<span data-ttu-id="809eb-480">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-481">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="809eb-483">optionalAttendees: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[](/javascript/api/outlook_1_6/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="809eb-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="809eb-484">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="809eb-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="809eb-485">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-486">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-486">Read mode</span></span>

<span data-ttu-id="809eb-487">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="809eb-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-488">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-488">Compose mode</span></span>

<span data-ttu-id="809eb-489">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="809eb-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="809eb-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-490">Type</span></span>

*   <span data-ttu-id="809eb-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="809eb-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-492">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-492">Requirements</span></span>

|<span data-ttu-id="809eb-493">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-493">Requirement</span></span>| <span data-ttu-id="809eb-494">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-495">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-496">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-496">1.0</span></span>|
|[<span data-ttu-id="809eb-497">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-498">ReadItem</span></span>|
|[<span data-ttu-id="809eb-499">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-500">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="809eb-501">organizador: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="809eb-501">organizer: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="809eb-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-504">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-504">Type</span></span>

*   [<span data-ttu-id="809eb-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="809eb-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="809eb-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-506">Requirements</span></span>

|<span data-ttu-id="809eb-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-507">Requirement</span></span>| <span data-ttu-id="809eb-508">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-510">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-510">1.0</span></span>|
|[<span data-ttu-id="809eb-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-512">ReadItem</span></span>|
|[<span data-ttu-id="809eb-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-514">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="809eb-516">requiredAttendees: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[](/javascript/api/outlook_1_6/office.recipients) de matriz. <</span><span class="sxs-lookup"><span data-stu-id="809eb-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="809eb-517">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="809eb-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="809eb-518">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-519">Read mode</span></span>

<span data-ttu-id="809eb-520">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="809eb-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-521">Compose mode</span></span>

<span data-ttu-id="809eb-522">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="809eb-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="809eb-523">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-523">Type</span></span>

*   <span data-ttu-id="809eb-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="809eb-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-525">Requirements</span></span>

|<span data-ttu-id="809eb-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-526">Requirement</span></span>| <span data-ttu-id="809eb-527">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-528">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-529">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-529">1.0</span></span>|
|[<span data-ttu-id="809eb-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-531">ReadItem</span></span>|
|[<span data-ttu-id="809eb-532">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-533">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="809eb-534">remetente: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="809eb-534">sender: [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="809eb-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="809eb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="809eb-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="809eb-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-539">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="809eb-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="809eb-540">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-540">Type</span></span>

*   [<span data-ttu-id="809eb-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="809eb-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="809eb-542">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-542">Requirements</span></span>

|<span data-ttu-id="809eb-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-543">Requirement</span></span>| <span data-ttu-id="809eb-544">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-546">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-546">1.0</span></span>|
|[<span data-ttu-id="809eb-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-548">ReadItem</span></span>|
|[<span data-ttu-id="809eb-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-550">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-551">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="809eb-552">Início: data | [Tempo](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="809eb-552">start: Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="809eb-553">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="809eb-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="809eb-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="809eb-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-556">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-556">Read mode</span></span>

<span data-ttu-id="809eb-557">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="809eb-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-558">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-558">Compose mode</span></span>

<span data-ttu-id="809eb-559">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="809eb-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="809eb-560">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="809eb-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="809eb-561">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="809eb-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="809eb-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-562">Type</span></span>

*   <span data-ttu-id="809eb-563">Data | [Hora](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="809eb-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-564">Requirements</span></span>

|<span data-ttu-id="809eb-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-565">Requirement</span></span>| <span data-ttu-id="809eb-566">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-568">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-568">1.0</span></span>|
|[<span data-ttu-id="809eb-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-570">ReadItem</span></span>|
|[<span data-ttu-id="809eb-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-572">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="809eb-573">subject: cadeia de caracteres | [Assunto](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="809eb-573">subject: String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="809eb-574">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="809eb-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="809eb-575">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="809eb-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-576">Read mode</span></span>

<span data-ttu-id="809eb-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="809eb-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-579">Compose mode</span></span>

<span data-ttu-id="809eb-580">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="809eb-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="809eb-581">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-581">Type</span></span>

*   <span data-ttu-id="809eb-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="809eb-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-583">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-583">Requirements</span></span>

|<span data-ttu-id="809eb-584">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-584">Requirement</span></span>| <span data-ttu-id="809eb-585">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-586">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-587">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-587">1.0</span></span>|
|[<span data-ttu-id="809eb-588">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-589">ReadItem</span></span>|
|[<span data-ttu-id="809eb-590">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-591">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-591">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="809eb-592">para: Array. <[](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[destinatários](/javascript/api/outlook_1_6/office.recipients) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="809eb-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="809eb-593">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="809eb-594">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="809eb-595">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="809eb-595">Read mode</span></span>

<span data-ttu-id="809eb-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="809eb-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="809eb-598">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="809eb-598">Compose mode</span></span>

<span data-ttu-id="809eb-599">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="809eb-600">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-600">Type</span></span>

*   <span data-ttu-id="809eb-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="809eb-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-602">Requirements</span></span>

|<span data-ttu-id="809eb-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-603">Requirement</span></span>| <span data-ttu-id="809eb-604">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-606">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-606">1.0</span></span>|
|[<span data-ttu-id="809eb-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-608">ReadItem</span></span>|
|[<span data-ttu-id="809eb-609">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-610">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="809eb-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="809eb-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="809eb-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="809eb-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="809eb-613">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="809eb-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="809eb-614">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="809eb-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="809eb-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="809eb-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-616">Parameters</span></span>

|<span data-ttu-id="809eb-617">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-617">Name</span></span>| <span data-ttu-id="809eb-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-618">Type</span></span>| <span data-ttu-id="809eb-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-619">Attributes</span></span>| <span data-ttu-id="809eb-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="809eb-621">String</span><span class="sxs-lookup"><span data-stu-id="809eb-621">String</span></span>||<span data-ttu-id="809eb-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="809eb-624">String</span><span class="sxs-lookup"><span data-stu-id="809eb-624">String</span></span>||<span data-ttu-id="809eb-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="809eb-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-627">Object</span></span>| <span data-ttu-id="809eb-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-628">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-629">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="809eb-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-630">Object</span></span> | <span data-ttu-id="809eb-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-631">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="809eb-633">Booliano</span><span class="sxs-lookup"><span data-stu-id="809eb-633">Boolean</span></span> | <span data-ttu-id="809eb-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-634">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-635">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="809eb-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="809eb-636">function</span><span class="sxs-lookup"><span data-stu-id="809eb-636">function</span></span>| <span data-ttu-id="809eb-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-637">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="809eb-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="809eb-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="809eb-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="809eb-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="809eb-641">Erros</span><span class="sxs-lookup"><span data-stu-id="809eb-641">Errors</span></span>

| <span data-ttu-id="809eb-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="809eb-642">Error code</span></span> | <span data-ttu-id="809eb-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="809eb-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="809eb-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="809eb-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="809eb-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="809eb-646">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="809eb-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-647">Requirements</span></span>

|<span data-ttu-id="809eb-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-648">Requirement</span></span>| <span data-ttu-id="809eb-649">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-651">1.1</span><span class="sxs-lookup"><span data-stu-id="809eb-651">1.1</span></span>|
|[<span data-ttu-id="809eb-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-655">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="809eb-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="809eb-656">Examples</span></span>

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

<span data-ttu-id="809eb-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="809eb-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="809eb-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="809eb-659">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="809eb-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="809eb-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="809eb-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="809eb-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="809eb-664">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="809eb-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-665">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-665">Parameters</span></span>

|<span data-ttu-id="809eb-666">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-666">Name</span></span>| <span data-ttu-id="809eb-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-667">Type</span></span>| <span data-ttu-id="809eb-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-668">Attributes</span></span>| <span data-ttu-id="809eb-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="809eb-670">String</span><span class="sxs-lookup"><span data-stu-id="809eb-670">String</span></span>||<span data-ttu-id="809eb-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="809eb-673">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="809eb-673">String</span></span>||<span data-ttu-id="809eb-674">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="809eb-674">The subject of the item to be attached.</span></span> <span data-ttu-id="809eb-675">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="809eb-676">Object</span><span class="sxs-lookup"><span data-stu-id="809eb-676">Object</span></span>| <span data-ttu-id="809eb-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-677">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="809eb-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-679">Object</span></span>| <span data-ttu-id="809eb-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-680">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="809eb-682">function</span><span class="sxs-lookup"><span data-stu-id="809eb-682">function</span></span>| <span data-ttu-id="809eb-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-683">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="809eb-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="809eb-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="809eb-686">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="809eb-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="809eb-687">Erros</span><span class="sxs-lookup"><span data-stu-id="809eb-687">Errors</span></span>

| <span data-ttu-id="809eb-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="809eb-688">Error code</span></span> | <span data-ttu-id="809eb-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="809eb-690">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="809eb-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-691">Requirements</span></span>

|<span data-ttu-id="809eb-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-692">Requirement</span></span>| <span data-ttu-id="809eb-693">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-695">1.1</span><span class="sxs-lookup"><span data-stu-id="809eb-695">1.1</span></span>|
|[<span data-ttu-id="809eb-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-699">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-700">Example</span></span>

<span data-ttu-id="809eb-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="809eb-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="809eb-702">close()</span><span class="sxs-lookup"><span data-stu-id="809eb-702">close()</span></span>

<span data-ttu-id="809eb-703">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="809eb-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="809eb-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="809eb-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-706">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="809eb-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="809eb-707">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="809eb-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-708">Requirements</span></span>

|<span data-ttu-id="809eb-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-709">Requirement</span></span>| <span data-ttu-id="809eb-710">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-712">1.3</span><span class="sxs-lookup"><span data-stu-id="809eb-712">1.3</span></span>|
|[<span data-ttu-id="809eb-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="809eb-714">Restricted</span></span>|
|[<span data-ttu-id="809eb-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-716">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="809eb-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="809eb-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="809eb-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="809eb-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-719">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-720">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="809eb-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="809eb-721">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="809eb-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="809eb-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="809eb-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-725">Parameters</span></span>

| <span data-ttu-id="809eb-726">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-726">Name</span></span> | <span data-ttu-id="809eb-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-727">Type</span></span> | <span data-ttu-id="809eb-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-728">Attributes</span></span> | <span data-ttu-id="809eb-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="809eb-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="809eb-730">String &#124; Object</span></span>| |<span data-ttu-id="809eb-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="809eb-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="809eb-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="809eb-733">**OR**</span></span><br/><span data-ttu-id="809eb-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="809eb-736">String</span><span class="sxs-lookup"><span data-stu-id="809eb-736">String</span></span> | <span data-ttu-id="809eb-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-737">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="809eb-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="809eb-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="809eb-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-741">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-742">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="809eb-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="809eb-743">String</span><span class="sxs-lookup"><span data-stu-id="809eb-743">String</span></span> | | <span data-ttu-id="809eb-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="809eb-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="809eb-746">String</span><span class="sxs-lookup"><span data-stu-id="809eb-746">String</span></span> | | <span data-ttu-id="809eb-747">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="809eb-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="809eb-748">String</span><span class="sxs-lookup"><span data-stu-id="809eb-748">String</span></span> | | <span data-ttu-id="809eb-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="809eb-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="809eb-751">Booliano</span><span class="sxs-lookup"><span data-stu-id="809eb-751">Boolean</span></span> | | <span data-ttu-id="809eb-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="809eb-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="809eb-754">String</span><span class="sxs-lookup"><span data-stu-id="809eb-754">String</span></span> | | <span data-ttu-id="809eb-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="809eb-758">function</span><span class="sxs-lookup"><span data-stu-id="809eb-758">function</span></span> | <span data-ttu-id="809eb-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-759">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-761">Requirements</span></span>

|<span data-ttu-id="809eb-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-762">Requirement</span></span>| <span data-ttu-id="809eb-763">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-765">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-765">1.0</span></span>|
|[<span data-ttu-id="809eb-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-767">ReadItem</span></span>|
|[<span data-ttu-id="809eb-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-769">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="809eb-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="809eb-770">Examples</span></span>

<span data-ttu-id="809eb-771">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="809eb-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="809eb-772">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="809eb-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="809eb-773">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="809eb-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="809eb-774">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="809eb-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="809eb-775">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="809eb-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="809eb-776">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="809eb-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="809eb-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="809eb-778">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="809eb-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-779">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-780">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="809eb-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="809eb-781">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="809eb-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="809eb-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="809eb-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-785">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-785">Parameters</span></span>

| <span data-ttu-id="809eb-786">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-786">Name</span></span> | <span data-ttu-id="809eb-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-787">Type</span></span> | <span data-ttu-id="809eb-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-788">Attributes</span></span> | <span data-ttu-id="809eb-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="809eb-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="809eb-790">String &#124; Object</span></span>| | <span data-ttu-id="809eb-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="809eb-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="809eb-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="809eb-793">**OR**</span></span><br/><span data-ttu-id="809eb-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="809eb-796">String</span><span class="sxs-lookup"><span data-stu-id="809eb-796">String</span></span> | <span data-ttu-id="809eb-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-797">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="809eb-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="809eb-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="809eb-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-801">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-802">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="809eb-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="809eb-803">String</span><span class="sxs-lookup"><span data-stu-id="809eb-803">String</span></span> | | <span data-ttu-id="809eb-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="809eb-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="809eb-806">String</span><span class="sxs-lookup"><span data-stu-id="809eb-806">String</span></span> | | <span data-ttu-id="809eb-807">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="809eb-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="809eb-808">String</span><span class="sxs-lookup"><span data-stu-id="809eb-808">String</span></span> | | <span data-ttu-id="809eb-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="809eb-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="809eb-811">Booliano</span><span class="sxs-lookup"><span data-stu-id="809eb-811">Boolean</span></span> | | <span data-ttu-id="809eb-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="809eb-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="809eb-814">String</span><span class="sxs-lookup"><span data-stu-id="809eb-814">String</span></span> | | <span data-ttu-id="809eb-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="809eb-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="809eb-818">function</span><span class="sxs-lookup"><span data-stu-id="809eb-818">function</span></span> | <span data-ttu-id="809eb-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-819">&lt;optional&gt;</span></span> | <span data-ttu-id="809eb-820">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-821">Requirements</span></span>

|<span data-ttu-id="809eb-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-822">Requirement</span></span>| <span data-ttu-id="809eb-823">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-824">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-825">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-825">1.0</span></span>|
|[<span data-ttu-id="809eb-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-827">ReadItem</span></span>|
|[<span data-ttu-id="809eb-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-829">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="809eb-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="809eb-830">Examples</span></span>

<span data-ttu-id="809eb-831">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="809eb-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="809eb-832">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="809eb-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="809eb-833">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="809eb-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="809eb-834">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="809eb-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="809eb-835">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="809eb-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="809eb-836">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="809eb-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="809eb-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="809eb-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="809eb-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-839">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-840">Requirements</span></span>

|<span data-ttu-id="809eb-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-841">Requirement</span></span>| <span data-ttu-id="809eb-842">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-844">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-844">1.0</span></span>|
|[<span data-ttu-id="809eb-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-846">ReadItem</span></span>|
|[<span data-ttu-id="809eb-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-848">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-849">Returns:</span></span>

<span data-ttu-id="809eb-850">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="809eb-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="809eb-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-851">Example</span></span>

<span data-ttu-id="809eb-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="809eb-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="809eb-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="809eb-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="809eb-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-855">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-856">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-856">Parameters</span></span>

|<span data-ttu-id="809eb-857">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-857">Name</span></span>| <span data-ttu-id="809eb-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-858">Type</span></span>| <span data-ttu-id="809eb-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="809eb-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="809eb-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="809eb-861">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="809eb-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-862">Requirements</span></span>

|<span data-ttu-id="809eb-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-863">Requirement</span></span>| <span data-ttu-id="809eb-864">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-865">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-866">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-866">1.0</span></span>|
|[<span data-ttu-id="809eb-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="809eb-868">Restricted</span></span>|
|[<span data-ttu-id="809eb-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-870">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-871">Returns:</span></span>

<span data-ttu-id="809eb-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="809eb-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="809eb-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="809eb-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="809eb-874">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="809eb-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="809eb-875">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="809eb-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="809eb-876">Value of `entityType`</span></span> | <span data-ttu-id="809eb-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="809eb-877">Type of objects in returned array</span></span> | <span data-ttu-id="809eb-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="809eb-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="809eb-879">String</span><span class="sxs-lookup"><span data-stu-id="809eb-879">String</span></span> | <span data-ttu-id="809eb-880">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="809eb-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="809eb-881">Contato</span><span class="sxs-lookup"><span data-stu-id="809eb-881">Contact</span></span> | <span data-ttu-id="809eb-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="809eb-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="809eb-883">String</span><span class="sxs-lookup"><span data-stu-id="809eb-883">String</span></span> | <span data-ttu-id="809eb-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="809eb-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="809eb-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="809eb-885">MeetingSuggestion</span></span> | <span data-ttu-id="809eb-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="809eb-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="809eb-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="809eb-887">PhoneNumber</span></span> | <span data-ttu-id="809eb-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="809eb-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="809eb-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="809eb-889">TaskSuggestion</span></span> | <span data-ttu-id="809eb-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="809eb-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="809eb-891">String</span><span class="sxs-lookup"><span data-stu-id="809eb-891">String</span></span> | <span data-ttu-id="809eb-892">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="809eb-892">**Restricted**</span></span> |

<span data-ttu-id="809eb-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="809eb-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="809eb-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-894">Example</span></span>

<span data-ttu-id="809eb-895">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="809eb-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="809eb-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="809eb-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="809eb-897">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="809eb-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-898">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="809eb-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-900">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-900">Parameters</span></span>

|<span data-ttu-id="809eb-901">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-901">Name</span></span>| <span data-ttu-id="809eb-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-902">Type</span></span>| <span data-ttu-id="809eb-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="809eb-904">String</span><span class="sxs-lookup"><span data-stu-id="809eb-904">String</span></span>|<span data-ttu-id="809eb-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="809eb-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-906">Requirements</span></span>

|<span data-ttu-id="809eb-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-907">Requirement</span></span>| <span data-ttu-id="809eb-908">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-909">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-910">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-910">1.0</span></span>|
|[<span data-ttu-id="809eb-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-912">ReadItem</span></span>|
|[<span data-ttu-id="809eb-913">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-914">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-915">Returns:</span></span>

<span data-ttu-id="809eb-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="809eb-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="809eb-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="809eb-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="809eb-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="809eb-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="809eb-920">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="809eb-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-921">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="809eb-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="809eb-925">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="809eb-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="809eb-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="809eb-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="809eb-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="809eb-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-930">Requirements</span></span>

|<span data-ttu-id="809eb-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-931">Requirement</span></span>| <span data-ttu-id="809eb-932">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-934">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-934">1.0</span></span>|
|[<span data-ttu-id="809eb-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-936">ReadItem</span></span>|
|[<span data-ttu-id="809eb-937">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-938">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-939">Returns:</span></span>

<span data-ttu-id="809eb-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="809eb-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="809eb-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="809eb-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="809eb-943">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="809eb-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-944">Example</span></span>

<span data-ttu-id="809eb-945">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="809eb-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="809eb-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="809eb-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="809eb-947">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="809eb-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-948">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-949">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="809eb-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="809eb-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="809eb-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-952">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-952">Parameters</span></span>

|<span data-ttu-id="809eb-953">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-953">Name</span></span>| <span data-ttu-id="809eb-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-954">Type</span></span>| <span data-ttu-id="809eb-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="809eb-956">String</span><span class="sxs-lookup"><span data-stu-id="809eb-956">String</span></span>|<span data-ttu-id="809eb-957">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="809eb-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-958">Requirements</span></span>

|<span data-ttu-id="809eb-959">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-959">Requirement</span></span>| <span data-ttu-id="809eb-960">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-961">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-962">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-962">1.0</span></span>|
|[<span data-ttu-id="809eb-963">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-964">ReadItem</span></span>|
|[<span data-ttu-id="809eb-965">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-966">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-967">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-967">Returns:</span></span>

<span data-ttu-id="809eb-968">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="809eb-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="809eb-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="809eb-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="809eb-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="809eb-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="809eb-971">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="809eb-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="809eb-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="809eb-973">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="809eb-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="809eb-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-976">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-976">Parameters</span></span>

|<span data-ttu-id="809eb-977">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-977">Name</span></span>| <span data-ttu-id="809eb-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-978">Type</span></span>| <span data-ttu-id="809eb-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-979">Attributes</span></span>| <span data-ttu-id="809eb-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="809eb-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="809eb-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="809eb-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="809eb-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="809eb-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-985">Object</span></span>| <span data-ttu-id="809eb-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-986">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-987">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="809eb-988">Object</span><span class="sxs-lookup"><span data-stu-id="809eb-988">Object</span></span>| <span data-ttu-id="809eb-989">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-989">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-990">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="809eb-991">function</span><span class="sxs-lookup"><span data-stu-id="809eb-991">function</span></span>||<span data-ttu-id="809eb-992">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="809eb-993">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="809eb-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="809eb-994">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="809eb-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-995">Requirements</span></span>

|<span data-ttu-id="809eb-996">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-996">Requirement</span></span>| <span data-ttu-id="809eb-997">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-998">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-999">1.2</span><span class="sxs-lookup"><span data-stu-id="809eb-999">1.2</span></span>|
|[<span data-ttu-id="809eb-1000">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-1002">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1003">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-1004">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-1004">Returns:</span></span>

<span data-ttu-id="809eb-1005">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="809eb-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="809eb-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="809eb-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="809eb-1007">String</span><span class="sxs-lookup"><span data-stu-id="809eb-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="809eb-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="809eb-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="809eb-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="809eb-1010">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="809eb-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="809eb-1011">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="809eb-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-1012">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1013">Requirements</span></span>

|<span data-ttu-id="809eb-1014">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1014">Requirement</span></span>| <span data-ttu-id="809eb-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1016">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="809eb-1017">1.6</span></span> |
|[<span data-ttu-id="809eb-1018">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1019">ReadItem</span></span>|
|[<span data-ttu-id="809eb-1020">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1021">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-1022">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-1022">Returns:</span></span>

<span data-ttu-id="809eb-1023">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="809eb-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="809eb-1024">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1024">Example</span></span>

<span data-ttu-id="809eb-1025">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="809eb-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="809eb-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="809eb-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="809eb-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="809eb-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-1029">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="809eb-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="809eb-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="809eb-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="809eb-1033">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="809eb-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="809eb-1034">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="809eb-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="809eb-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="809eb-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="809eb-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1038">Requirements</span></span>

|<span data-ttu-id="809eb-1039">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1039">Requirement</span></span>| <span data-ttu-id="809eb-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1041">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="809eb-1042">1.6</span></span> |
|[<span data-ttu-id="809eb-1043">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1044">ReadItem</span></span>|
|[<span data-ttu-id="809eb-1045">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1046">Read</span><span class="sxs-lookup"><span data-stu-id="809eb-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="809eb-1047">Retorna:</span><span class="sxs-lookup"><span data-stu-id="809eb-1047">Returns:</span></span>

<span data-ttu-id="809eb-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="809eb-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="809eb-1050">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1050">Example</span></span>

<span data-ttu-id="809eb-1051">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="809eb-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="809eb-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="809eb-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="809eb-1053">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="809eb-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="809eb-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="809eb-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-1057">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-1057">Parameters</span></span>

|<span data-ttu-id="809eb-1058">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-1058">Name</span></span>| <span data-ttu-id="809eb-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-1059">Type</span></span>| <span data-ttu-id="809eb-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-1060">Attributes</span></span>| <span data-ttu-id="809eb-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="809eb-1062">function</span><span class="sxs-lookup"><span data-stu-id="809eb-1062">function</span></span>||<span data-ttu-id="809eb-1063">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="809eb-1064">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="809eb-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="809eb-1065">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="809eb-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="809eb-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1066">Object</span></span>| <span data-ttu-id="809eb-1067">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1068">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="809eb-1069">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1070">Requirements</span></span>

|<span data-ttu-id="809eb-1071">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1071">Requirement</span></span>| <span data-ttu-id="809eb-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1073">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="809eb-1074">1.0</span></span>|
|[<span data-ttu-id="809eb-1075">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1076">ReadItem</span></span>|
|[<span data-ttu-id="809eb-1077">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="809eb-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1078">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="809eb-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1079">Example</span></span>

<span data-ttu-id="809eb-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="809eb-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="809eb-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="809eb-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="809eb-1084">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="809eb-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="809eb-p172">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="809eb-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-1089">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-1089">Parameters</span></span>

|<span data-ttu-id="809eb-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-1090">Name</span></span>| <span data-ttu-id="809eb-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-1091">Type</span></span>| <span data-ttu-id="809eb-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-1092">Attributes</span></span>| <span data-ttu-id="809eb-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="809eb-1094">String</span><span class="sxs-lookup"><span data-stu-id="809eb-1094">String</span></span>||<span data-ttu-id="809eb-1095">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="809eb-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="809eb-1096">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1096">Object</span></span>| <span data-ttu-id="809eb-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1098">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="809eb-1099">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1099">Object</span></span>| <span data-ttu-id="809eb-1100">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1101">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="809eb-1102">function</span><span class="sxs-lookup"><span data-stu-id="809eb-1102">function</span></span>| <span data-ttu-id="809eb-1103">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1104">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="809eb-1105">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="809eb-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="809eb-1106">Erros</span><span class="sxs-lookup"><span data-stu-id="809eb-1106">Errors</span></span>

| <span data-ttu-id="809eb-1107">Código de erro</span><span class="sxs-lookup"><span data-stu-id="809eb-1107">Error code</span></span> | <span data-ttu-id="809eb-1108">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="809eb-1109">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="809eb-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-1110">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1110">Requirements</span></span>

|<span data-ttu-id="809eb-1111">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1111">Requirement</span></span>| <span data-ttu-id="809eb-1112">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1113">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="809eb-1114">1.1</span></span>|
|[<span data-ttu-id="809eb-1115">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-1117">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1118">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-1119">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1119">Example</span></span>

<span data-ttu-id="809eb-1120">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="809eb-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="809eb-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="809eb-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="809eb-1122">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="809eb-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="809eb-p173">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="809eb-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-1126">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="809eb-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="809eb-1127">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="809eb-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="809eb-p175">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="809eb-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="809eb-1131">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="809eb-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="809eb-1132">O Outlook para Mac não dá suporte à gravação de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="809eb-1132">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="809eb-1133">O `saveAsync` método falha quando chamado de uma reunião no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="809eb-1133">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="809eb-1134">Consulte [não é possível salvar uma reunião como rascunho no Outlook para Mac usando a API do Office js](https://support.microsoft.com/help/4505745) para obter uma solução alternativa.</span><span class="sxs-lookup"><span data-stu-id="809eb-1134">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="809eb-1135">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="809eb-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-1136">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-1136">Parameters</span></span>

|<span data-ttu-id="809eb-1137">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-1137">Name</span></span>| <span data-ttu-id="809eb-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-1138">Type</span></span>| <span data-ttu-id="809eb-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-1139">Attributes</span></span>| <span data-ttu-id="809eb-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="809eb-1141">Object</span><span class="sxs-lookup"><span data-stu-id="809eb-1141">Object</span></span>| <span data-ttu-id="809eb-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1143">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="809eb-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1144">Object</span></span>| <span data-ttu-id="809eb-1145">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1146">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="809eb-1147">function</span><span class="sxs-lookup"><span data-stu-id="809eb-1147">function</span></span>||<span data-ttu-id="809eb-1148">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="809eb-1149">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="809eb-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="809eb-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1150">Requirements</span></span>

|<span data-ttu-id="809eb-1151">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1151">Requirement</span></span>| <span data-ttu-id="809eb-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1153">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="809eb-1154">1.3</span></span>|
|[<span data-ttu-id="809eb-1155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-1157">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-1157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1158">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="809eb-1159">Exemplos</span><span class="sxs-lookup"><span data-stu-id="809eb-1159">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="809eb-p177">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="809eb-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="809eb-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="809eb-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="809eb-1163">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="809eb-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="809eb-p178">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="809eb-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="809eb-1167">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="809eb-1167">Parameters</span></span>

|<span data-ttu-id="809eb-1168">Nome</span><span class="sxs-lookup"><span data-stu-id="809eb-1168">Name</span></span>| <span data-ttu-id="809eb-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="809eb-1169">Type</span></span>| <span data-ttu-id="809eb-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="809eb-1170">Attributes</span></span>| <span data-ttu-id="809eb-1171">Descrição</span><span class="sxs-lookup"><span data-stu-id="809eb-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="809eb-1172">String</span><span class="sxs-lookup"><span data-stu-id="809eb-1172">String</span></span>||<span data-ttu-id="809eb-p179">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="809eb-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="809eb-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1176">Object</span></span>| <span data-ttu-id="809eb-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1178">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="809eb-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="809eb-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="809eb-1179">Object</span></span>| <span data-ttu-id="809eb-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-1181">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="809eb-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="809eb-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="809eb-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="809eb-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="809eb-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="809eb-p180">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="809eb-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="809eb-p181">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="809eb-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="809eb-1188">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="809eb-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="809eb-1189">function</span><span class="sxs-lookup"><span data-stu-id="809eb-1189">function</span></span>||<span data-ttu-id="809eb-1190">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="809eb-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="809eb-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="809eb-1191">Requirements</span></span>

|<span data-ttu-id="809eb-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="809eb-1192">Requirement</span></span>| <span data-ttu-id="809eb-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="809eb-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="809eb-1194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="809eb-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="809eb-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="809eb-1195">1.2</span></span>|
|[<span data-ttu-id="809eb-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="809eb-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="809eb-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="809eb-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="809eb-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="809eb-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="809eb-1199">Escrever</span><span class="sxs-lookup"><span data-stu-id="809eb-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="809eb-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="809eb-1200">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
