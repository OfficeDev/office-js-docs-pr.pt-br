---
title: Office. Context. Mailbox. Item-conjunto de requisitos 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 009adf0730edc0e619a9fe15f20af07246da39a4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871135"
---
# <a name="item"></a><span data-ttu-id="281ba-102">item</span><span class="sxs-lookup"><span data-stu-id="281ba-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="281ba-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="281ba-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="281ba-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="281ba-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-106">Requirements</span></span>

|<span data-ttu-id="281ba-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-107">Requirement</span></span>| <span data-ttu-id="281ba-108">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-110">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-110">1.0</span></span>|
|[<span data-ttu-id="281ba-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="281ba-112">Restricted</span></span>|
|[<span data-ttu-id="281ba-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="281ba-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="281ba-115">Members and methods</span></span>

| <span data-ttu-id="281ba-116">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-116">Member</span></span> | <span data-ttu-id="281ba-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="281ba-118">attachments</span><span class="sxs-lookup"><span data-stu-id="281ba-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="281ba-119">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-119">Member</span></span> |
| [<span data-ttu-id="281ba-120">bcc</span><span class="sxs-lookup"><span data-stu-id="281ba-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="281ba-121">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-121">Member</span></span> |
| [<span data-ttu-id="281ba-122">body</span><span class="sxs-lookup"><span data-stu-id="281ba-122">body</span></span>](#body-body) | <span data-ttu-id="281ba-123">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-123">Member</span></span> |
| [<span data-ttu-id="281ba-124">cc</span><span class="sxs-lookup"><span data-stu-id="281ba-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="281ba-125">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-125">Member</span></span> |
| [<span data-ttu-id="281ba-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="281ba-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="281ba-127">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-127">Member</span></span> |
| [<span data-ttu-id="281ba-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="281ba-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="281ba-129">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-129">Member</span></span> |
| [<span data-ttu-id="281ba-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="281ba-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="281ba-131">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-131">Member</span></span> |
| [<span data-ttu-id="281ba-132">end</span><span class="sxs-lookup"><span data-stu-id="281ba-132">end</span></span>](#end-datetime) | <span data-ttu-id="281ba-133">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-133">Member</span></span> |
| [<span data-ttu-id="281ba-134">from</span><span class="sxs-lookup"><span data-stu-id="281ba-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="281ba-135">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-135">Member</span></span> |
| [<span data-ttu-id="281ba-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="281ba-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="281ba-137">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-137">Member</span></span> |
| [<span data-ttu-id="281ba-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="281ba-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="281ba-139">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-139">Member</span></span> |
| [<span data-ttu-id="281ba-140">itemId</span><span class="sxs-lookup"><span data-stu-id="281ba-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="281ba-141">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-141">Member</span></span> |
| [<span data-ttu-id="281ba-142">itemType</span><span class="sxs-lookup"><span data-stu-id="281ba-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="281ba-143">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-143">Member</span></span> |
| [<span data-ttu-id="281ba-144">location</span><span class="sxs-lookup"><span data-stu-id="281ba-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="281ba-145">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-145">Member</span></span> |
| [<span data-ttu-id="281ba-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="281ba-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="281ba-147">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-147">Member</span></span> |
| [<span data-ttu-id="281ba-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="281ba-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="281ba-149">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-149">Member</span></span> |
| [<span data-ttu-id="281ba-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="281ba-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="281ba-151">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-151">Member</span></span> |
| [<span data-ttu-id="281ba-152">organizer</span><span class="sxs-lookup"><span data-stu-id="281ba-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="281ba-153">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-153">Member</span></span> |
| [<span data-ttu-id="281ba-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="281ba-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="281ba-155">Member</span><span class="sxs-lookup"><span data-stu-id="281ba-155">Member</span></span> |
| [<span data-ttu-id="281ba-156">sender</span><span class="sxs-lookup"><span data-stu-id="281ba-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="281ba-157">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-157">Member</span></span> |
| [<span data-ttu-id="281ba-158">start</span><span class="sxs-lookup"><span data-stu-id="281ba-158">start</span></span>](#start-datetime) | <span data-ttu-id="281ba-159">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-159">Member</span></span> |
| [<span data-ttu-id="281ba-160">subject</span><span class="sxs-lookup"><span data-stu-id="281ba-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="281ba-161">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-161">Member</span></span> |
| [<span data-ttu-id="281ba-162">to</span><span class="sxs-lookup"><span data-stu-id="281ba-162">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="281ba-163">Membro</span><span class="sxs-lookup"><span data-stu-id="281ba-163">Member</span></span> |
| [<span data-ttu-id="281ba-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="281ba-165">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-165">Method</span></span> |
| [<span data-ttu-id="281ba-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="281ba-167">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-167">Method</span></span> |
| [<span data-ttu-id="281ba-168">close</span><span class="sxs-lookup"><span data-stu-id="281ba-168">close</span></span>](#close) | <span data-ttu-id="281ba-169">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-169">Method</span></span> |
| [<span data-ttu-id="281ba-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="281ba-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="281ba-171">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-171">Method</span></span> |
| [<span data-ttu-id="281ba-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="281ba-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="281ba-173">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-173">Method</span></span> |
| [<span data-ttu-id="281ba-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="281ba-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="281ba-175">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-175">Method</span></span> |
| [<span data-ttu-id="281ba-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="281ba-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="281ba-177">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-177">Method</span></span> |
| [<span data-ttu-id="281ba-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="281ba-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="281ba-179">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-179">Method</span></span> |
| [<span data-ttu-id="281ba-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="281ba-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="281ba-181">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-181">Method</span></span> |
| [<span data-ttu-id="281ba-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="281ba-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="281ba-183">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-183">Method</span></span> |
| [<span data-ttu-id="281ba-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="281ba-185">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-185">Method</span></span> |
| [<span data-ttu-id="281ba-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="281ba-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="281ba-187">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-187">Method</span></span> |
| [<span data-ttu-id="281ba-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="281ba-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="281ba-189">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-189">Method</span></span> |
| [<span data-ttu-id="281ba-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="281ba-191">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-191">Method</span></span> |
| [<span data-ttu-id="281ba-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="281ba-193">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-193">Method</span></span> |
| [<span data-ttu-id="281ba-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="281ba-195">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-195">Method</span></span> |
| [<span data-ttu-id="281ba-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="281ba-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="281ba-197">Método</span><span class="sxs-lookup"><span data-stu-id="281ba-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="281ba-198">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-198">Example</span></span>

<span data-ttu-id="281ba-199">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="281ba-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="281ba-200">Membros</span><span class="sxs-lookup"><span data-stu-id="281ba-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="281ba-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="281ba-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="281ba-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-204">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="281ba-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="281ba-205">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="281ba-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-206">Type</span></span>

*   <span data-ttu-id="281ba-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="281ba-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-208">Requirements</span></span>

|<span data-ttu-id="281ba-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-209">Requirement</span></span>| <span data-ttu-id="281ba-210">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-212">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-212">1.0</span></span>|
|[<span data-ttu-id="281ba-213">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-214">ReadItem</span></span>|
|[<span data-ttu-id="281ba-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-216">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-217">Example</span></span>

<span data-ttu-id="281ba-218">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="281ba-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="281ba-220">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="281ba-221">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="281ba-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-222">Type</span></span>

*   [<span data-ttu-id="281ba-223">Destinatários</span><span class="sxs-lookup"><span data-stu-id="281ba-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="281ba-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-224">Requirements</span></span>

|<span data-ttu-id="281ba-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-225">Requirement</span></span>| <span data-ttu-id="281ba-226">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-228">1.1</span><span class="sxs-lookup"><span data-stu-id="281ba-228">1.1</span></span>|
|[<span data-ttu-id="281ba-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-230">ReadItem</span></span>|
|[<span data-ttu-id="281ba-231">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-232">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-233">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="281ba-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="281ba-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="281ba-235">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="281ba-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-236">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-236">Type</span></span>

*   [<span data-ttu-id="281ba-237">Body</span><span class="sxs-lookup"><span data-stu-id="281ba-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="281ba-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-238">Requirements</span></span>

|<span data-ttu-id="281ba-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-239">Requirement</span></span>| <span data-ttu-id="281ba-240">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-242">1.1</span><span class="sxs-lookup"><span data-stu-id="281ba-242">1.1</span></span>|
|[<span data-ttu-id="281ba-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-244">ReadItem</span></span>|
|[<span data-ttu-id="281ba-245">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-247">Example</span></span>

<span data-ttu-id="281ba-248">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="281ba-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="281ba-249">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="281ba-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="281ba-251">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="281ba-252">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-253">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-253">Read mode</span></span>

<span data-ttu-id="281ba-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="281ba-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-256">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-256">Compose mode</span></span>

<span data-ttu-id="281ba-257">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="281ba-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-258">Type</span></span>

*   <span data-ttu-id="281ba-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-260">Requirements</span></span>

|<span data-ttu-id="281ba-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-261">Requirement</span></span>| <span data-ttu-id="281ba-262">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-264">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-264">1.0</span></span>|
|[<span data-ttu-id="281ba-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-266">ReadItem</span></span>|
|[<span data-ttu-id="281ba-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="281ba-269">(anulável) conversationId :Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="281ba-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="281ba-270">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="281ba-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="281ba-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="281ba-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="281ba-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="281ba-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-275">Type</span></span>

*   <span data-ttu-id="281ba-276">String</span><span class="sxs-lookup"><span data-stu-id="281ba-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-277">Requirements</span></span>

|<span data-ttu-id="281ba-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-278">Requirement</span></span>| <span data-ttu-id="281ba-279">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-281">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-281">1.0</span></span>|
|[<span data-ttu-id="281ba-282">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-283">ReadItem</span></span>|
|[<span data-ttu-id="281ba-284">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-286">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="281ba-287">dateTimeCreated :Data</span><span class="sxs-lookup"><span data-stu-id="281ba-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="281ba-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-290">Type</span></span>

*   <span data-ttu-id="281ba-291">Data</span><span class="sxs-lookup"><span data-stu-id="281ba-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-292">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-292">Requirements</span></span>

|<span data-ttu-id="281ba-293">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-293">Requirement</span></span>| <span data-ttu-id="281ba-294">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-295">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-296">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-296">1.0</span></span>|
|[<span data-ttu-id="281ba-297">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-298">ReadItem</span></span>|
|[<span data-ttu-id="281ba-299">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-300">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-301">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="281ba-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="281ba-302">dateTimeModified :Date</span></span>

<span data-ttu-id="281ba-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-305">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-306">Type</span></span>

*   <span data-ttu-id="281ba-307">Data</span><span class="sxs-lookup"><span data-stu-id="281ba-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-308">Requirements</span></span>

|<span data-ttu-id="281ba-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-309">Requirement</span></span>| <span data-ttu-id="281ba-310">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-312">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-312">1.0</span></span>|
|[<span data-ttu-id="281ba-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-314">ReadItem</span></span>|
|[<span data-ttu-id="281ba-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-316">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="281ba-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="281ba-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="281ba-319">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="281ba-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="281ba-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="281ba-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-322">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-322">Read mode</span></span>

<span data-ttu-id="281ba-323">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="281ba-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-324">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-324">Compose mode</span></span>

<span data-ttu-id="281ba-325">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="281ba-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="281ba-326">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="281ba-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="281ba-327">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="281ba-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="281ba-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-328">Type</span></span>

*   <span data-ttu-id="281ba-329">Data | [Hora](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="281ba-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-330">Requirements</span></span>

|<span data-ttu-id="281ba-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-331">Requirement</span></span>| <span data-ttu-id="281ba-332">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-334">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-334">1.0</span></span>|
|[<span data-ttu-id="281ba-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-336">ReadItem</span></span>|
|[<span data-ttu-id="281ba-337">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-338">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="281ba-339">De:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="281ba-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="281ba-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="281ba-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="281ba-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-344">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="281ba-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-345">Type</span></span>

*   [<span data-ttu-id="281ba-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="281ba-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="281ba-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="281ba-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-348">Requirements</span></span>

|<span data-ttu-id="281ba-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-349">Requirement</span></span>| <span data-ttu-id="281ba-350">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-352">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-352">1.0</span></span>|
|[<span data-ttu-id="281ba-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-354">ReadItem</span></span>|
|[<span data-ttu-id="281ba-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-356">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="281ba-357">internetMessageId Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="281ba-357">internetMessageId :String</span></span>

<span data-ttu-id="281ba-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-360">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-360">Type</span></span>

*   <span data-ttu-id="281ba-361">String</span><span class="sxs-lookup"><span data-stu-id="281ba-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-362">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-362">Requirements</span></span>

|<span data-ttu-id="281ba-363">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-363">Requirement</span></span>| <span data-ttu-id="281ba-364">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-365">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-366">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-366">1.0</span></span>|
|[<span data-ttu-id="281ba-367">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-368">ReadItem</span></span>|
|[<span data-ttu-id="281ba-369">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-370">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-371">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="281ba-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="281ba-372">itemClass :String</span></span>

<span data-ttu-id="281ba-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="281ba-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="281ba-377">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-377">Type</span></span> | <span data-ttu-id="281ba-378">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-378">Description</span></span> | <span data-ttu-id="281ba-379">classe de item</span><span class="sxs-lookup"><span data-stu-id="281ba-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="281ba-380">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="281ba-380">Appointment items</span></span> | <span data-ttu-id="281ba-381">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="281ba-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="281ba-382">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="281ba-382">Message items</span></span> | <span data-ttu-id="281ba-383">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="281ba-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="281ba-384">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="281ba-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-385">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-385">Type</span></span>

*   <span data-ttu-id="281ba-386">String</span><span class="sxs-lookup"><span data-stu-id="281ba-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-387">Requirements</span></span>

|<span data-ttu-id="281ba-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-388">Requirement</span></span>| <span data-ttu-id="281ba-389">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-391">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-391">1.0</span></span>|
|[<span data-ttu-id="281ba-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-393">ReadItem</span></span>|
|[<span data-ttu-id="281ba-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-395">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-396">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="281ba-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="281ba-397">(nullable) itemId :String</span></span>

<span data-ttu-id="281ba-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-400">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="281ba-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="281ba-401">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="281ba-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="281ba-402">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="281ba-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="281ba-403">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="281ba-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="281ba-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-406">Type</span></span>

*   <span data-ttu-id="281ba-407">String</span><span class="sxs-lookup"><span data-stu-id="281ba-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-408">Requirements</span></span>

|<span data-ttu-id="281ba-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-409">Requirement</span></span>| <span data-ttu-id="281ba-410">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-412">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-412">1.0</span></span>|
|[<span data-ttu-id="281ba-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-414">ReadItem</span></span>|
|[<span data-ttu-id="281ba-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-416">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-417">Example</span></span>

<span data-ttu-id="281ba-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="281ba-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="281ba-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="281ba-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="281ba-421">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="281ba-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="281ba-422">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-423">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-423">Type</span></span>

*   [<span data-ttu-id="281ba-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="281ba-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="281ba-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-425">Requirements</span></span>

|<span data-ttu-id="281ba-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-426">Requirement</span></span>| <span data-ttu-id="281ba-427">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-429">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-429">1.0</span></span>|
|[<span data-ttu-id="281ba-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-431">ReadItem</span></span>|
|[<span data-ttu-id="281ba-432">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-433">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="281ba-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="281ba-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="281ba-436">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-437">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-437">Read mode</span></span>

<span data-ttu-id="281ba-438">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-439">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-439">Compose mode</span></span>

<span data-ttu-id="281ba-440">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="281ba-441">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-441">Type</span></span>

*   <span data-ttu-id="281ba-442">Cadeia de caracteres | [Localização](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="281ba-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-443">Requirements</span></span>

|<span data-ttu-id="281ba-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-444">Requirement</span></span>| <span data-ttu-id="281ba-445">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-446">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-447">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-447">1.0</span></span>|
|[<span data-ttu-id="281ba-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-449">ReadItem</span></span>|
|[<span data-ttu-id="281ba-450">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-451">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="281ba-452">normalizedSubject :Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="281ba-452">normalizedSubject :String</span></span>

<span data-ttu-id="281ba-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="281ba-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="281ba-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-457">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-457">Type</span></span>

*   <span data-ttu-id="281ba-458">String</span><span class="sxs-lookup"><span data-stu-id="281ba-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-459">Requirements</span></span>

|<span data-ttu-id="281ba-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-460">Requirement</span></span>| <span data-ttu-id="281ba-461">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-463">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-463">1.0</span></span>|
|[<span data-ttu-id="281ba-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-465">ReadItem</span></span>|
|[<span data-ttu-id="281ba-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-467">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="281ba-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="281ba-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="281ba-470">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="281ba-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-471">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-471">Type</span></span>

*   [<span data-ttu-id="281ba-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="281ba-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="281ba-473">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-473">Requirements</span></span>

|<span data-ttu-id="281ba-474">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-474">Requirement</span></span>| <span data-ttu-id="281ba-475">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-476">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-477">1.3</span><span class="sxs-lookup"><span data-stu-id="281ba-477">1.3</span></span>|
|[<span data-ttu-id="281ba-478">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-479">ReadItem</span></span>|
|[<span data-ttu-id="281ba-480">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-481">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-482">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="281ba-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="281ba-484">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="281ba-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="281ba-485">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-486">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-486">Read mode</span></span>

<span data-ttu-id="281ba-487">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="281ba-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-488">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-488">Compose mode</span></span>

<span data-ttu-id="281ba-489">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="281ba-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="281ba-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-490">Type</span></span>

*   <span data-ttu-id="281ba-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-492">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-492">Requirements</span></span>

|<span data-ttu-id="281ba-493">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-493">Requirement</span></span>| <span data-ttu-id="281ba-494">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-495">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-496">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-496">1.0</span></span>|
|[<span data-ttu-id="281ba-497">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-498">ReadItem</span></span>|
|[<span data-ttu-id="281ba-499">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-500">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="281ba-501">organizador:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="281ba-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="281ba-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-504">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-504">Type</span></span>

*   [<span data-ttu-id="281ba-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="281ba-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="281ba-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-506">Requirements</span></span>

|<span data-ttu-id="281ba-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-507">Requirement</span></span>| <span data-ttu-id="281ba-508">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-510">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-510">1.0</span></span>|
|[<span data-ttu-id="281ba-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-512">ReadItem</span></span>|
|[<span data-ttu-id="281ba-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-514">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="281ba-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="281ba-517">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="281ba-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="281ba-518">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-519">Read mode</span></span>

<span data-ttu-id="281ba-520">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="281ba-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-521">Compose mode</span></span>

<span data-ttu-id="281ba-522">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="281ba-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="281ba-523">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-523">Type</span></span>

*   <span data-ttu-id="281ba-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-525">Requirements</span></span>

|<span data-ttu-id="281ba-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-526">Requirement</span></span>| <span data-ttu-id="281ba-527">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-528">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-529">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-529">1.0</span></span>|
|[<span data-ttu-id="281ba-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-531">ReadItem</span></span>|
|[<span data-ttu-id="281ba-532">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-533">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="281ba-534">remetente :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="281ba-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="281ba-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="281ba-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="281ba-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="281ba-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-539">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="281ba-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="281ba-540">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-540">Type</span></span>

*   [<span data-ttu-id="281ba-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="281ba-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="281ba-542">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-542">Requirements</span></span>

|<span data-ttu-id="281ba-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-543">Requirement</span></span>| <span data-ttu-id="281ba-544">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-545">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-546">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-546">1.0</span></span>|
|[<span data-ttu-id="281ba-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-548">ReadItem</span></span>|
|[<span data-ttu-id="281ba-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-550">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-551">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="281ba-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="281ba-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="281ba-553">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="281ba-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="281ba-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="281ba-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-556">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-556">Read mode</span></span>

<span data-ttu-id="281ba-557">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="281ba-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-558">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-558">Compose mode</span></span>

<span data-ttu-id="281ba-559">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="281ba-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="281ba-560">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="281ba-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="281ba-561">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="281ba-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="281ba-562">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-562">Type</span></span>

*   <span data-ttu-id="281ba-563">Data | [Hora](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="281ba-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-564">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-564">Requirements</span></span>

|<span data-ttu-id="281ba-565">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-565">Requirement</span></span>| <span data-ttu-id="281ba-566">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-567">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-568">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-568">1.0</span></span>|
|[<span data-ttu-id="281ba-569">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-570">ReadItem</span></span>|
|[<span data-ttu-id="281ba-571">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-572">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="281ba-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="281ba-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="281ba-574">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="281ba-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="281ba-575">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="281ba-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-576">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-576">Read mode</span></span>

<span data-ttu-id="281ba-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="281ba-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-579">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-579">Compose mode</span></span>

<span data-ttu-id="281ba-580">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="281ba-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="281ba-581">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-581">Type</span></span>

*   <span data-ttu-id="281ba-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="281ba-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-583">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-583">Requirements</span></span>

|<span data-ttu-id="281ba-584">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-584">Requirement</span></span>| <span data-ttu-id="281ba-585">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-586">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-587">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-587">1.0</span></span>|
|[<span data-ttu-id="281ba-588">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-589">ReadItem</span></span>|
|[<span data-ttu-id="281ba-590">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-591">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="281ba-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="281ba-593">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="281ba-594">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="281ba-595">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="281ba-595">Read mode</span></span>

<span data-ttu-id="281ba-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="281ba-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="281ba-598">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="281ba-598">Compose mode</span></span>

<span data-ttu-id="281ba-599">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="281ba-600">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-600">Type</span></span>

*   <span data-ttu-id="281ba-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="281ba-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-602">Requirements</span></span>

|<span data-ttu-id="281ba-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-603">Requirement</span></span>| <span data-ttu-id="281ba-604">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-606">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-606">1.0</span></span>|
|[<span data-ttu-id="281ba-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-608">ReadItem</span></span>|
|[<span data-ttu-id="281ba-609">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-610">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="281ba-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="281ba-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="281ba-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="281ba-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="281ba-613">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="281ba-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="281ba-614">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="281ba-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="281ba-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="281ba-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-616">Parameters</span></span>

|<span data-ttu-id="281ba-617">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-617">Name</span></span>| <span data-ttu-id="281ba-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-618">Type</span></span>| <span data-ttu-id="281ba-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-619">Attributes</span></span>| <span data-ttu-id="281ba-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="281ba-621">String</span><span class="sxs-lookup"><span data-stu-id="281ba-621">String</span></span>||<span data-ttu-id="281ba-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="281ba-624">String</span><span class="sxs-lookup"><span data-stu-id="281ba-624">String</span></span>||<span data-ttu-id="281ba-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="281ba-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-627">Object</span></span>| <span data-ttu-id="281ba-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-628">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-629">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="281ba-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-630">Object</span></span> | <span data-ttu-id="281ba-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-631">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="281ba-633">Booliano</span><span class="sxs-lookup"><span data-stu-id="281ba-633">Boolean</span></span> | <span data-ttu-id="281ba-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-634">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-635">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="281ba-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="281ba-636">function</span><span class="sxs-lookup"><span data-stu-id="281ba-636">function</span></span>| <span data-ttu-id="281ba-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-637">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="281ba-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="281ba-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="281ba-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="281ba-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="281ba-641">Erros</span><span class="sxs-lookup"><span data-stu-id="281ba-641">Errors</span></span>

| <span data-ttu-id="281ba-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="281ba-642">Error code</span></span> | <span data-ttu-id="281ba-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="281ba-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="281ba-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="281ba-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="281ba-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="281ba-646">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="281ba-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-647">Requirements</span></span>

|<span data-ttu-id="281ba-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-648">Requirement</span></span>| <span data-ttu-id="281ba-649">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-651">1.1</span><span class="sxs-lookup"><span data-stu-id="281ba-651">1.1</span></span>|
|[<span data-ttu-id="281ba-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-655">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="281ba-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="281ba-656">Examples</span></span>

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

<span data-ttu-id="281ba-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="281ba-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="281ba-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="281ba-659">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="281ba-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="281ba-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="281ba-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="281ba-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="281ba-664">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="281ba-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-665">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-665">Parameters</span></span>

|<span data-ttu-id="281ba-666">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-666">Name</span></span>| <span data-ttu-id="281ba-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-667">Type</span></span>| <span data-ttu-id="281ba-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-668">Attributes</span></span>| <span data-ttu-id="281ba-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="281ba-670">String</span><span class="sxs-lookup"><span data-stu-id="281ba-670">String</span></span>||<span data-ttu-id="281ba-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="281ba-673">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="281ba-673">String</span></span>||<span data-ttu-id="281ba-674">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="281ba-674">The subject of the item to be attached.</span></span> <span data-ttu-id="281ba-675">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="281ba-676">Object</span><span class="sxs-lookup"><span data-stu-id="281ba-676">Object</span></span>| <span data-ttu-id="281ba-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-677">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="281ba-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-679">Object</span></span>| <span data-ttu-id="281ba-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-680">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="281ba-682">function</span><span class="sxs-lookup"><span data-stu-id="281ba-682">function</span></span>| <span data-ttu-id="281ba-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-683">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="281ba-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="281ba-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="281ba-686">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="281ba-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="281ba-687">Erros</span><span class="sxs-lookup"><span data-stu-id="281ba-687">Errors</span></span>

| <span data-ttu-id="281ba-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="281ba-688">Error code</span></span> | <span data-ttu-id="281ba-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="281ba-690">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="281ba-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-691">Requirements</span></span>

|<span data-ttu-id="281ba-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-692">Requirement</span></span>| <span data-ttu-id="281ba-693">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-695">1.1</span><span class="sxs-lookup"><span data-stu-id="281ba-695">1.1</span></span>|
|[<span data-ttu-id="281ba-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-699">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-700">Example</span></span>

<span data-ttu-id="281ba-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="281ba-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="281ba-702">close()</span><span class="sxs-lookup"><span data-stu-id="281ba-702">close()</span></span>

<span data-ttu-id="281ba-703">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="281ba-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="281ba-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="281ba-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-706">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="281ba-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="281ba-707">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="281ba-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-708">Requirements</span></span>

|<span data-ttu-id="281ba-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-709">Requirement</span></span>| <span data-ttu-id="281ba-710">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-712">1.3</span><span class="sxs-lookup"><span data-stu-id="281ba-712">1.3</span></span>|
|[<span data-ttu-id="281ba-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="281ba-714">Restricted</span></span>|
|[<span data-ttu-id="281ba-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-716">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="281ba-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="281ba-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="281ba-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="281ba-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-719">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-720">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="281ba-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="281ba-721">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="281ba-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="281ba-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="281ba-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-725">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-725">Parameters</span></span>

| <span data-ttu-id="281ba-726">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-726">Name</span></span> | <span data-ttu-id="281ba-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-727">Type</span></span> | <span data-ttu-id="281ba-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-728">Attributes</span></span> | <span data-ttu-id="281ba-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="281ba-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="281ba-730">String &#124; Object</span></span>| |<span data-ttu-id="281ba-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="281ba-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="281ba-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="281ba-733">**OR**</span></span><br/><span data-ttu-id="281ba-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="281ba-736">String</span><span class="sxs-lookup"><span data-stu-id="281ba-736">String</span></span> | <span data-ttu-id="281ba-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-737">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="281ba-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="281ba-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="281ba-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-741">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-742">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="281ba-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="281ba-743">String</span><span class="sxs-lookup"><span data-stu-id="281ba-743">String</span></span> | | <span data-ttu-id="281ba-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="281ba-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="281ba-746">String</span><span class="sxs-lookup"><span data-stu-id="281ba-746">String</span></span> | | <span data-ttu-id="281ba-747">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="281ba-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="281ba-748">String</span><span class="sxs-lookup"><span data-stu-id="281ba-748">String</span></span> | | <span data-ttu-id="281ba-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="281ba-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="281ba-751">Booliano</span><span class="sxs-lookup"><span data-stu-id="281ba-751">Boolean</span></span> | | <span data-ttu-id="281ba-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="281ba-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="281ba-754">String</span><span class="sxs-lookup"><span data-stu-id="281ba-754">String</span></span> | | <span data-ttu-id="281ba-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="281ba-758">function</span><span class="sxs-lookup"><span data-stu-id="281ba-758">function</span></span> | <span data-ttu-id="281ba-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-759">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-761">Requirements</span></span>

|<span data-ttu-id="281ba-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-762">Requirement</span></span>| <span data-ttu-id="281ba-763">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-765">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-765">1.0</span></span>|
|[<span data-ttu-id="281ba-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-767">ReadItem</span></span>|
|[<span data-ttu-id="281ba-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-769">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="281ba-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="281ba-770">Examples</span></span>

<span data-ttu-id="281ba-771">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="281ba-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="281ba-772">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="281ba-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="281ba-773">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="281ba-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="281ba-774">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="281ba-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="281ba-775">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="281ba-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="281ba-776">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="281ba-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="281ba-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="281ba-778">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="281ba-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-779">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-780">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="281ba-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="281ba-781">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="281ba-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="281ba-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="281ba-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-785">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-785">Parameters</span></span>

| <span data-ttu-id="281ba-786">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-786">Name</span></span> | <span data-ttu-id="281ba-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-787">Type</span></span> | <span data-ttu-id="281ba-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-788">Attributes</span></span> | <span data-ttu-id="281ba-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="281ba-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="281ba-790">String &#124; Object</span></span>| | <span data-ttu-id="281ba-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="281ba-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="281ba-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="281ba-793">**OR**</span></span><br/><span data-ttu-id="281ba-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="281ba-796">String</span><span class="sxs-lookup"><span data-stu-id="281ba-796">String</span></span> | <span data-ttu-id="281ba-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-797">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="281ba-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="281ba-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="281ba-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-801">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-802">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="281ba-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="281ba-803">String</span><span class="sxs-lookup"><span data-stu-id="281ba-803">String</span></span> | | <span data-ttu-id="281ba-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="281ba-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="281ba-806">String</span><span class="sxs-lookup"><span data-stu-id="281ba-806">String</span></span> | | <span data-ttu-id="281ba-807">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="281ba-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="281ba-808">String</span><span class="sxs-lookup"><span data-stu-id="281ba-808">String</span></span> | | <span data-ttu-id="281ba-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="281ba-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="281ba-811">Booliano</span><span class="sxs-lookup"><span data-stu-id="281ba-811">Boolean</span></span> | | <span data-ttu-id="281ba-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="281ba-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="281ba-814">String</span><span class="sxs-lookup"><span data-stu-id="281ba-814">String</span></span> | | <span data-ttu-id="281ba-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="281ba-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="281ba-818">function</span><span class="sxs-lookup"><span data-stu-id="281ba-818">function</span></span> | <span data-ttu-id="281ba-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-819">&lt;optional&gt;</span></span> | <span data-ttu-id="281ba-820">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-821">Requirements</span></span>

|<span data-ttu-id="281ba-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-822">Requirement</span></span>| <span data-ttu-id="281ba-823">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-824">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-825">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-825">1.0</span></span>|
|[<span data-ttu-id="281ba-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-827">ReadItem</span></span>|
|[<span data-ttu-id="281ba-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-829">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="281ba-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="281ba-830">Examples</span></span>

<span data-ttu-id="281ba-831">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="281ba-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="281ba-832">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="281ba-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="281ba-833">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="281ba-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="281ba-834">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="281ba-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="281ba-835">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="281ba-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="281ba-836">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="281ba-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="281ba-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="281ba-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="281ba-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-839">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-840">Requirements</span></span>

|<span data-ttu-id="281ba-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-841">Requirement</span></span>| <span data-ttu-id="281ba-842">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-844">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-844">1.0</span></span>|
|[<span data-ttu-id="281ba-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-846">ReadItem</span></span>|
|[<span data-ttu-id="281ba-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-848">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-849">Returns:</span></span>

<span data-ttu-id="281ba-850">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="281ba-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="281ba-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-851">Example</span></span>

<span data-ttu-id="281ba-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="281ba-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="281ba-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="281ba-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="281ba-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-855">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-856">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-856">Parameters</span></span>

|<span data-ttu-id="281ba-857">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-857">Name</span></span>| <span data-ttu-id="281ba-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-858">Type</span></span>| <span data-ttu-id="281ba-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="281ba-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="281ba-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="281ba-861">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="281ba-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-862">Requirements</span></span>

|<span data-ttu-id="281ba-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-863">Requirement</span></span>| <span data-ttu-id="281ba-864">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-865">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-866">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-866">1.0</span></span>|
|[<span data-ttu-id="281ba-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="281ba-868">Restricted</span></span>|
|[<span data-ttu-id="281ba-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-870">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-871">Returns:</span></span>

<span data-ttu-id="281ba-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="281ba-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="281ba-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="281ba-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="281ba-874">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="281ba-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="281ba-875">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="281ba-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="281ba-876">Value of `entityType`</span></span> | <span data-ttu-id="281ba-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="281ba-877">Type of objects in returned array</span></span> | <span data-ttu-id="281ba-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="281ba-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="281ba-879">String</span><span class="sxs-lookup"><span data-stu-id="281ba-879">String</span></span> | <span data-ttu-id="281ba-880">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="281ba-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="281ba-881">Contato</span><span class="sxs-lookup"><span data-stu-id="281ba-881">Contact</span></span> | <span data-ttu-id="281ba-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="281ba-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="281ba-883">String</span><span class="sxs-lookup"><span data-stu-id="281ba-883">String</span></span> | <span data-ttu-id="281ba-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="281ba-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="281ba-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="281ba-885">MeetingSuggestion</span></span> | <span data-ttu-id="281ba-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="281ba-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="281ba-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="281ba-887">PhoneNumber</span></span> | <span data-ttu-id="281ba-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="281ba-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="281ba-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="281ba-889">TaskSuggestion</span></span> | <span data-ttu-id="281ba-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="281ba-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="281ba-891">String</span><span class="sxs-lookup"><span data-stu-id="281ba-891">String</span></span> | <span data-ttu-id="281ba-892">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="281ba-892">**Restricted**</span></span> |

<span data-ttu-id="281ba-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="281ba-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="281ba-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-894">Example</span></span>

<span data-ttu-id="281ba-895">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="281ba-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="281ba-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="281ba-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="281ba-897">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="281ba-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-898">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="281ba-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-900">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-900">Parameters</span></span>

|<span data-ttu-id="281ba-901">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-901">Name</span></span>| <span data-ttu-id="281ba-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-902">Type</span></span>| <span data-ttu-id="281ba-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="281ba-904">String</span><span class="sxs-lookup"><span data-stu-id="281ba-904">String</span></span>|<span data-ttu-id="281ba-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="281ba-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-906">Requirements</span></span>

|<span data-ttu-id="281ba-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-907">Requirement</span></span>| <span data-ttu-id="281ba-908">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-909">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-910">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-910">1.0</span></span>|
|[<span data-ttu-id="281ba-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-912">ReadItem</span></span>|
|[<span data-ttu-id="281ba-913">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-914">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-915">Returns:</span></span>

<span data-ttu-id="281ba-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="281ba-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="281ba-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="281ba-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="281ba-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="281ba-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="281ba-920">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="281ba-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-921">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="281ba-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="281ba-925">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="281ba-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="281ba-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="281ba-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="281ba-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="281ba-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-930">Requirements</span></span>

|<span data-ttu-id="281ba-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-931">Requirement</span></span>| <span data-ttu-id="281ba-932">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-934">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-934">1.0</span></span>|
|[<span data-ttu-id="281ba-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-936">ReadItem</span></span>|
|[<span data-ttu-id="281ba-937">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-938">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-939">Returns:</span></span>

<span data-ttu-id="281ba-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="281ba-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="281ba-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="281ba-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="281ba-943">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="281ba-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-944">Example</span></span>

<span data-ttu-id="281ba-945">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="281ba-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="281ba-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="281ba-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="281ba-947">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="281ba-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-948">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-949">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="281ba-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="281ba-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="281ba-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-952">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-952">Parameters</span></span>

|<span data-ttu-id="281ba-953">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-953">Name</span></span>| <span data-ttu-id="281ba-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-954">Type</span></span>| <span data-ttu-id="281ba-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="281ba-956">String</span><span class="sxs-lookup"><span data-stu-id="281ba-956">String</span></span>|<span data-ttu-id="281ba-957">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="281ba-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-958">Requirements</span></span>

|<span data-ttu-id="281ba-959">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-959">Requirement</span></span>| <span data-ttu-id="281ba-960">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-961">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-962">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-962">1.0</span></span>|
|[<span data-ttu-id="281ba-963">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-964">ReadItem</span></span>|
|[<span data-ttu-id="281ba-965">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-966">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-967">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-967">Returns:</span></span>

<span data-ttu-id="281ba-968">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="281ba-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="281ba-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="281ba-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="281ba-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="281ba-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="281ba-971">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="281ba-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="281ba-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="281ba-973">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="281ba-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="281ba-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-976">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-976">Parameters</span></span>

|<span data-ttu-id="281ba-977">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-977">Name</span></span>| <span data-ttu-id="281ba-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-978">Type</span></span>| <span data-ttu-id="281ba-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-979">Attributes</span></span>| <span data-ttu-id="281ba-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="281ba-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="281ba-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="281ba-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="281ba-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="281ba-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-985">Object</span></span>| <span data-ttu-id="281ba-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-986">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-987">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="281ba-988">Object</span><span class="sxs-lookup"><span data-stu-id="281ba-988">Object</span></span>| <span data-ttu-id="281ba-989">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-989">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-990">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="281ba-991">function</span><span class="sxs-lookup"><span data-stu-id="281ba-991">function</span></span>||<span data-ttu-id="281ba-992">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="281ba-993">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="281ba-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="281ba-994">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="281ba-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-995">Requirements</span></span>

|<span data-ttu-id="281ba-996">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-996">Requirement</span></span>| <span data-ttu-id="281ba-997">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-998">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-999">1.2</span><span class="sxs-lookup"><span data-stu-id="281ba-999">1.2</span></span>|
|[<span data-ttu-id="281ba-1000">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-1002">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1003">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-1004">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-1004">Returns:</span></span>

<span data-ttu-id="281ba-1005">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="281ba-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="281ba-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="281ba-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="281ba-1007">String</span><span class="sxs-lookup"><span data-stu-id="281ba-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="281ba-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="281ba-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="281ba-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="281ba-1010">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou.</span><span class="sxs-lookup"><span data-stu-id="281ba-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="281ba-1011">As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="281ba-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-1012">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1013">Requirements</span></span>

|<span data-ttu-id="281ba-1014">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1014">Requirement</span></span>| <span data-ttu-id="281ba-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1016">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="281ba-1017">1.6</span></span> |
|[<span data-ttu-id="281ba-1018">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1019">ReadItem</span></span>|
|[<span data-ttu-id="281ba-1020">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1021">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-1022">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-1022">Returns:</span></span>

<span data-ttu-id="281ba-1023">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="281ba-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="281ba-1024">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1024">Example</span></span>

<span data-ttu-id="281ba-1025">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="281ba-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="281ba-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="281ba-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="281ba-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="281ba-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-1029">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="281ba-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="281ba-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="281ba-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="281ba-1033">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="281ba-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="281ba-1034">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="281ba-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="281ba-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="281ba-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="281ba-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1038">Requirements</span></span>

|<span data-ttu-id="281ba-1039">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1039">Requirement</span></span>| <span data-ttu-id="281ba-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1041">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="281ba-1042">1.6</span></span> |
|[<span data-ttu-id="281ba-1043">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1044">ReadItem</span></span>|
|[<span data-ttu-id="281ba-1045">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1046">Read</span><span class="sxs-lookup"><span data-stu-id="281ba-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="281ba-1047">Retorna:</span><span class="sxs-lookup"><span data-stu-id="281ba-1047">Returns:</span></span>

<span data-ttu-id="281ba-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="281ba-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="281ba-1050">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1050">Example</span></span>

<span data-ttu-id="281ba-1051">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="281ba-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="281ba-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="281ba-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="281ba-1053">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="281ba-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="281ba-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="281ba-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-1057">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-1057">Parameters</span></span>

|<span data-ttu-id="281ba-1058">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-1058">Name</span></span>| <span data-ttu-id="281ba-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-1059">Type</span></span>| <span data-ttu-id="281ba-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-1060">Attributes</span></span>| <span data-ttu-id="281ba-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="281ba-1062">function</span><span class="sxs-lookup"><span data-stu-id="281ba-1062">function</span></span>||<span data-ttu-id="281ba-1063">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="281ba-1064">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="281ba-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="281ba-1065">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="281ba-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="281ba-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1066">Object</span></span>| <span data-ttu-id="281ba-1067">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1068">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="281ba-1069">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1070">Requirements</span></span>

|<span data-ttu-id="281ba-1071">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1071">Requirement</span></span>| <span data-ttu-id="281ba-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1073">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="281ba-1074">1.0</span></span>|
|[<span data-ttu-id="281ba-1075">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1076">ReadItem</span></span>|
|[<span data-ttu-id="281ba-1077">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="281ba-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1078">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="281ba-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1079">Example</span></span>

<span data-ttu-id="281ba-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="281ba-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="281ba-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="281ba-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="281ba-1084">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="281ba-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="281ba-p172">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="281ba-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-1089">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-1089">Parameters</span></span>

|<span data-ttu-id="281ba-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-1090">Name</span></span>| <span data-ttu-id="281ba-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-1091">Type</span></span>| <span data-ttu-id="281ba-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-1092">Attributes</span></span>| <span data-ttu-id="281ba-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="281ba-1094">String</span><span class="sxs-lookup"><span data-stu-id="281ba-1094">String</span></span>||<span data-ttu-id="281ba-1095">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="281ba-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="281ba-1096">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1096">Object</span></span>| <span data-ttu-id="281ba-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1098">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="281ba-1099">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1099">Object</span></span>| <span data-ttu-id="281ba-1100">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1101">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="281ba-1102">function</span><span class="sxs-lookup"><span data-stu-id="281ba-1102">function</span></span>| <span data-ttu-id="281ba-1103">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1104">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="281ba-1105">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="281ba-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="281ba-1106">Erros</span><span class="sxs-lookup"><span data-stu-id="281ba-1106">Errors</span></span>

| <span data-ttu-id="281ba-1107">Código de erro</span><span class="sxs-lookup"><span data-stu-id="281ba-1107">Error code</span></span> | <span data-ttu-id="281ba-1108">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="281ba-1109">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="281ba-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-1110">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1110">Requirements</span></span>

|<span data-ttu-id="281ba-1111">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1111">Requirement</span></span>| <span data-ttu-id="281ba-1112">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1113">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="281ba-1114">1.1</span></span>|
|[<span data-ttu-id="281ba-1115">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-1117">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1118">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-1119">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1119">Example</span></span>

<span data-ttu-id="281ba-1120">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="281ba-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="281ba-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="281ba-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="281ba-1122">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="281ba-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="281ba-p173">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="281ba-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-1126">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="281ba-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="281ba-1127">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="281ba-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="281ba-p175">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="281ba-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="281ba-1131">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="281ba-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="281ba-1132">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="281ba-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="281ba-1133">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="281ba-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="281ba-1134">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="281ba-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-1135">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-1135">Parameters</span></span>

|<span data-ttu-id="281ba-1136">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-1136">Name</span></span>| <span data-ttu-id="281ba-1137">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-1137">Type</span></span>| <span data-ttu-id="281ba-1138">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-1138">Attributes</span></span>| <span data-ttu-id="281ba-1139">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="281ba-1140">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1140">Object</span></span>| <span data-ttu-id="281ba-1141">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1142">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="281ba-1143">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1143">Object</span></span>| <span data-ttu-id="281ba-1144">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1145">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="281ba-1146">function</span><span class="sxs-lookup"><span data-stu-id="281ba-1146">function</span></span>||<span data-ttu-id="281ba-1147">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="281ba-1148">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="281ba-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="281ba-1149">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1149">Requirements</span></span>

|<span data-ttu-id="281ba-1150">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1150">Requirement</span></span>| <span data-ttu-id="281ba-1151">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1152">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="281ba-1153">1.3</span></span>|
|[<span data-ttu-id="281ba-1154">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1154">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-1156">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-1156">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1157">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="281ba-1158">Exemplos</span><span class="sxs-lookup"><span data-stu-id="281ba-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="281ba-p177">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="281ba-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="281ba-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="281ba-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="281ba-1162">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="281ba-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="281ba-p178">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="281ba-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="281ba-1166">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="281ba-1166">Parameters</span></span>

|<span data-ttu-id="281ba-1167">Nome</span><span class="sxs-lookup"><span data-stu-id="281ba-1167">Name</span></span>| <span data-ttu-id="281ba-1168">Tipo</span><span class="sxs-lookup"><span data-stu-id="281ba-1168">Type</span></span>| <span data-ttu-id="281ba-1169">Atributos</span><span class="sxs-lookup"><span data-stu-id="281ba-1169">Attributes</span></span>| <span data-ttu-id="281ba-1170">Descrição</span><span class="sxs-lookup"><span data-stu-id="281ba-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="281ba-1171">String</span><span class="sxs-lookup"><span data-stu-id="281ba-1171">String</span></span>||<span data-ttu-id="281ba-p179">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="281ba-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="281ba-1175">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1175">Object</span></span>| <span data-ttu-id="281ba-1176">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1177">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="281ba-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="281ba-1178">Objeto</span><span class="sxs-lookup"><span data-stu-id="281ba-1178">Object</span></span>| <span data-ttu-id="281ba-1179">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-1180">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="281ba-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="281ba-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="281ba-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="281ba-1182">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="281ba-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="281ba-p180">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="281ba-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="281ba-p181">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="281ba-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="281ba-1187">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="281ba-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="281ba-1188">function</span><span class="sxs-lookup"><span data-stu-id="281ba-1188">function</span></span>||<span data-ttu-id="281ba-1189">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="281ba-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="281ba-1190">Requisitos</span><span class="sxs-lookup"><span data-stu-id="281ba-1190">Requirements</span></span>

|<span data-ttu-id="281ba-1191">Requisito</span><span class="sxs-lookup"><span data-stu-id="281ba-1191">Requirement</span></span>| <span data-ttu-id="281ba-1192">Valor</span><span class="sxs-lookup"><span data-stu-id="281ba-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="281ba-1193">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="281ba-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="281ba-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="281ba-1194">1.2</span></span>|
|[<span data-ttu-id="281ba-1195">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="281ba-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="281ba-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="281ba-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="281ba-1197">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="281ba-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="281ba-1198">Escrever</span><span class="sxs-lookup"><span data-stu-id="281ba-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="281ba-1199">Exemplo</span><span class="sxs-lookup"><span data-stu-id="281ba-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
