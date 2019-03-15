---
title: Office.context.mailbox.item - conjunto de requisitos 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Priority
ms.openlocfilehash: 999a9b732a7efd3f4a9c289e5ff0b97e87278e60
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600274"
---
# <a name="item"></a><span data-ttu-id="43c0f-102">item</span><span class="sxs-lookup"><span data-stu-id="43c0f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="43c0f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="43c0f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="43c0f-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="43c0f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-106">Requirements</span></span>

|<span data-ttu-id="43c0f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-107">Requirement</span></span>| <span data-ttu-id="43c0f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-110">1.0</span></span>|
|[<span data-ttu-id="43c0f-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="43c0f-112">Restricted</span></span>|
|[<span data-ttu-id="43c0f-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-114">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="43c0f-115">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="43c0f-115">Members and methods</span></span>

| <span data-ttu-id="43c0f-116">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-116">Member</span></span> | <span data-ttu-id="43c0f-117">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="43c0f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="43c0f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="43c0f-119">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-119">Member</span></span> |
| [<span data-ttu-id="43c0f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="43c0f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="43c0f-121">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-121">Member</span></span> |
| [<span data-ttu-id="43c0f-122">body</span><span class="sxs-lookup"><span data-stu-id="43c0f-122">body</span></span>](#body-body) | <span data-ttu-id="43c0f-123">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-123">Member</span></span> |
| [<span data-ttu-id="43c0f-124">cc</span><span class="sxs-lookup"><span data-stu-id="43c0f-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="43c0f-125">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-125">Member</span></span> |
| [<span data-ttu-id="43c0f-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="43c0f-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="43c0f-127">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-127">Member</span></span> |
| [<span data-ttu-id="43c0f-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="43c0f-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="43c0f-129">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-129">Member</span></span> |
| [<span data-ttu-id="43c0f-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="43c0f-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="43c0f-131">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-131">Member</span></span> |
| [<span data-ttu-id="43c0f-132">end</span><span class="sxs-lookup"><span data-stu-id="43c0f-132">end</span></span>](#end-datetime) | <span data-ttu-id="43c0f-133">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-133">Member</span></span> |
| [<span data-ttu-id="43c0f-134">from</span><span class="sxs-lookup"><span data-stu-id="43c0f-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="43c0f-135">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-135">Member</span></span> |
| [<span data-ttu-id="43c0f-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="43c0f-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="43c0f-137">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-137">Member</span></span> |
| [<span data-ttu-id="43c0f-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="43c0f-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="43c0f-139">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-139">Member</span></span> |
| [<span data-ttu-id="43c0f-140">itemId</span><span class="sxs-lookup"><span data-stu-id="43c0f-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="43c0f-141">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-141">Member</span></span> |
| [<span data-ttu-id="43c0f-142">itemType</span><span class="sxs-lookup"><span data-stu-id="43c0f-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="43c0f-143">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-143">Member</span></span> |
| [<span data-ttu-id="43c0f-144">location</span><span class="sxs-lookup"><span data-stu-id="43c0f-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="43c0f-145">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-145">Member</span></span> |
| [<span data-ttu-id="43c0f-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="43c0f-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="43c0f-147">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-147">Member</span></span> |
| [<span data-ttu-id="43c0f-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="43c0f-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="43c0f-149">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-149">Member</span></span> |
| [<span data-ttu-id="43c0f-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="43c0f-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="43c0f-151">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-151">Member</span></span> |
| [<span data-ttu-id="43c0f-152">organizer</span><span class="sxs-lookup"><span data-stu-id="43c0f-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="43c0f-153">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-153">Member</span></span> |
| [<span data-ttu-id="43c0f-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="43c0f-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="43c0f-155">Member</span><span class="sxs-lookup"><span data-stu-id="43c0f-155">Member</span></span> |
| [<span data-ttu-id="43c0f-156">sender</span><span class="sxs-lookup"><span data-stu-id="43c0f-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="43c0f-157">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-157">Member</span></span> |
| [<span data-ttu-id="43c0f-158">start</span><span class="sxs-lookup"><span data-stu-id="43c0f-158">start</span></span>](#start-datetime) | <span data-ttu-id="43c0f-159">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-159">Member</span></span> |
| [<span data-ttu-id="43c0f-160">subject</span><span class="sxs-lookup"><span data-stu-id="43c0f-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="43c0f-161">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-161">Member</span></span> |
| [<span data-ttu-id="43c0f-162">to</span><span class="sxs-lookup"><span data-stu-id="43c0f-162">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="43c0f-163">Membro</span><span class="sxs-lookup"><span data-stu-id="43c0f-163">Member</span></span> |
| [<span data-ttu-id="43c0f-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="43c0f-165">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-165">Method</span></span> |
| [<span data-ttu-id="43c0f-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="43c0f-167">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-167">Method</span></span> |
| [<span data-ttu-id="43c0f-168">close</span><span class="sxs-lookup"><span data-stu-id="43c0f-168">close</span></span>](#close) | <span data-ttu-id="43c0f-169">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-169">Method</span></span> |
| [<span data-ttu-id="43c0f-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="43c0f-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="43c0f-171">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-171">Method</span></span> |
| [<span data-ttu-id="43c0f-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="43c0f-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="43c0f-173">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-173">Method</span></span> |
| [<span data-ttu-id="43c0f-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="43c0f-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="43c0f-175">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-175">Method</span></span> |
| [<span data-ttu-id="43c0f-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="43c0f-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="43c0f-177">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-177">Method</span></span> |
| [<span data-ttu-id="43c0f-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="43c0f-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="43c0f-179">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-179">Method</span></span> |
| [<span data-ttu-id="43c0f-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="43c0f-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="43c0f-181">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-181">Method</span></span> |
| [<span data-ttu-id="43c0f-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="43c0f-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="43c0f-183">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-183">Method</span></span> |
| [<span data-ttu-id="43c0f-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="43c0f-185">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-185">Method</span></span> |
| [<span data-ttu-id="43c0f-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="43c0f-187">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-187">Method</span></span> |
| [<span data-ttu-id="43c0f-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="43c0f-189">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-189">Method</span></span> |
| [<span data-ttu-id="43c0f-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="43c0f-191">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-191">Method</span></span> |
| [<span data-ttu-id="43c0f-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="43c0f-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="43c0f-193">Método</span><span class="sxs-lookup"><span data-stu-id="43c0f-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="43c0f-194">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-194">Example</span></span>

<span data-ttu-id="43c0f-195">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="43c0f-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="43c0f-196">Membros</span><span class="sxs-lookup"><span data-stu-id="43c0f-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="43c0f-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="43c0f-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="43c0f-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-200">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="43c0f-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="43c0f-201">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="43c0f-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-202">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-202">Type</span></span>

*   <span data-ttu-id="43c0f-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="43c0f-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-204">Requirements</span></span>

|<span data-ttu-id="43c0f-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-205">Requirement</span></span>| <span data-ttu-id="43c0f-206">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-208">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-208">1.0</span></span>|
|[<span data-ttu-id="43c0f-209">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-210">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-212">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-213">Example</span></span>

<span data-ttu-id="43c0f-214">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="43c0f-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="43c0f-216">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="43c0f-217">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="43c0f-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-218">Type</span></span>

*   [<span data-ttu-id="43c0f-219">Destinatários</span><span class="sxs-lookup"><span data-stu-id="43c0f-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="43c0f-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-220">Requirements</span></span>

|<span data-ttu-id="43c0f-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-221">Requirement</span></span>| <span data-ttu-id="43c0f-222">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-224">1.1</span><span class="sxs-lookup"><span data-stu-id="43c0f-224">1.1</span></span>|
|[<span data-ttu-id="43c0f-225">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-226">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-228">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="43c0f-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="43c0f-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="43c0f-231">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-232">Type</span></span>

*   [<span data-ttu-id="43c0f-233">Body</span><span class="sxs-lookup"><span data-stu-id="43c0f-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="43c0f-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-234">Requirements</span></span>

|<span data-ttu-id="43c0f-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-235">Requirement</span></span>| <span data-ttu-id="43c0f-236">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-238">1.1</span><span class="sxs-lookup"><span data-stu-id="43c0f-238">1.1</span></span>|
|[<span data-ttu-id="43c0f-239">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-240">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-241">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-243">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-243">Example</span></span>

<span data-ttu-id="43c0f-244">Este exemplo obtém o corpo da mensagem em texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="43c0f-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="43c0f-245">A seguir apresentamos um exemplo do resultado do parâmetro passado à função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="43c0f-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="43c0f-247">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="43c0f-248">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-249">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-249">Read mode</span></span>

<span data-ttu-id="43c0f-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-252">Compose mode</span></span>

<span data-ttu-id="43c0f-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-254">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-254">Type</span></span>

*   <span data-ttu-id="43c0f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-256">Requirements</span></span>

|<span data-ttu-id="43c0f-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-257">Requirement</span></span>| <span data-ttu-id="43c0f-258">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-260">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-260">1.0</span></span>|
|[<span data-ttu-id="43c0f-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-262">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-263">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-264">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="43c0f-265">(anulável) conversationId :Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="43c0f-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="43c0f-266">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="43c0f-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="43c0f-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="43c0f-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-271">Type</span></span>

*   <span data-ttu-id="43c0f-272">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-273">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-273">Requirements</span></span>

|<span data-ttu-id="43c0f-274">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-274">Requirement</span></span>| <span data-ttu-id="43c0f-275">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-276">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-277">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-277">1.0</span></span>|
|[<span data-ttu-id="43c0f-278">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-278">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-279">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-280">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-280">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-281">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-282">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="43c0f-283">dateTimeCreated :Data</span><span class="sxs-lookup"><span data-stu-id="43c0f-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="43c0f-p109">Obtém a data e a hora em que um item foi criado. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-286">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-286">Type</span></span>

*   <span data-ttu-id="43c0f-287">Data</span><span class="sxs-lookup"><span data-stu-id="43c0f-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-288">Requirements</span></span>

|<span data-ttu-id="43c0f-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-289">Requirement</span></span>| <span data-ttu-id="43c0f-290">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-292">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-292">1.0</span></span>|
|[<span data-ttu-id="43c0f-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-294">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-296">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="43c0f-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="43c0f-298">dateTimeModified :Date</span></span>

<span data-ttu-id="43c0f-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-301">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-302">Type</span></span>

*   <span data-ttu-id="43c0f-303">Data</span><span class="sxs-lookup"><span data-stu-id="43c0f-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-304">Requirements</span></span>

|<span data-ttu-id="43c0f-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-305">Requirement</span></span>| <span data-ttu-id="43c0f-306">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-308">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-308">1.0</span></span>|
|[<span data-ttu-id="43c0f-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-310">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-312">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="43c0f-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="43c0f-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="43c0f-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="43c0f-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="43c0f-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-318">Read mode</span></span>

<span data-ttu-id="43c0f-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-320">Compose mode</span></span>

<span data-ttu-id="43c0f-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="43c0f-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="43c0f-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="43c0f-323">O exemplo a seguir define a hora de término de um compromisso usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="43c0f-324">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-324">Type</span></span>

*   <span data-ttu-id="43c0f-325">Data | [Hora](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="43c0f-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-326">Requirements</span></span>

|<span data-ttu-id="43c0f-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-327">Requirement</span></span>| <span data-ttu-id="43c0f-328">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-330">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-330">1.0</span></span>|
|[<span data-ttu-id="43c0f-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-332">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-333">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="43c0f-335">De:[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="43c0f-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="43c0f-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="43c0f-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-340">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-341">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-341">Type</span></span>

*   [<span data-ttu-id="43c0f-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43c0f-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="43c0f-343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-343">Requirements</span></span>

|<span data-ttu-id="43c0f-344">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-344">Requirement</span></span>| <span data-ttu-id="43c0f-345">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-346">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-347">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-347">1.0</span></span>|
|[<span data-ttu-id="43c0f-348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-349">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-350">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-351">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-352">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="43c0f-353">internetMessageId Cadeia de Caracteres</span><span class="sxs-lookup"><span data-stu-id="43c0f-353">internetMessageId :String</span></span>

<span data-ttu-id="43c0f-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-356">Type</span></span>

*   <span data-ttu-id="43c0f-357">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-358">Requirements</span></span>

|<span data-ttu-id="43c0f-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-359">Requirement</span></span>| <span data-ttu-id="43c0f-360">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-362">1.0</span></span>|
|[<span data-ttu-id="43c0f-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-364">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-366">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="43c0f-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="43c0f-368">itemClass :String</span></span>

<span data-ttu-id="43c0f-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="43c0f-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="43c0f-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-373">Type</span></span> | <span data-ttu-id="43c0f-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-374">Description</span></span> | <span data-ttu-id="43c0f-375">classe de item</span><span class="sxs-lookup"><span data-stu-id="43c0f-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="43c0f-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="43c0f-376">Appointment items</span></span> | <span data-ttu-id="43c0f-377">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="43c0f-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="43c0f-378">Message items</span></span> | <span data-ttu-id="43c0f-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="43c0f-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="43c0f-380">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-381">Type</span></span>

*   <span data-ttu-id="43c0f-382">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-383">Requirements</span></span>

|<span data-ttu-id="43c0f-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-384">Requirement</span></span>| <span data-ttu-id="43c0f-385">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-387">1.0</span></span>|
|[<span data-ttu-id="43c0f-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-389">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-391">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="43c0f-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="43c0f-393">(nullable) itemId :String</span></span>

<span data-ttu-id="43c0f-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="43c0f-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="43c0f-397">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="43c0f-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="43c0f-398">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="43c0f-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="43c0f-399">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="43c0f-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="43c0f-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-402">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-402">Type</span></span>

*   <span data-ttu-id="43c0f-403">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-404">Requirements</span></span>

|<span data-ttu-id="43c0f-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-405">Requirement</span></span>| <span data-ttu-id="43c0f-406">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-407">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-408">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-408">1.0</span></span>|
|[<span data-ttu-id="43c0f-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-410">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-412">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-413">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-413">Example</span></span>

<span data-ttu-id="43c0f-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="43c0f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="43c0f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="43c0f-417">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="43c0f-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="43c0f-418">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-419">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-419">Type</span></span>

*   [<span data-ttu-id="43c0f-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="43c0f-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="43c0f-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-421">Requirements</span></span>

|<span data-ttu-id="43c0f-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-422">Requirement</span></span>| <span data-ttu-id="43c0f-423">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-425">1.0</span></span>|
|[<span data-ttu-id="43c0f-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-427">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-428">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-429">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="43c0f-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="43c0f-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="43c0f-432">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-433">Read mode</span></span>

<span data-ttu-id="43c0f-434">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-435">Compose mode</span></span>

<span data-ttu-id="43c0f-436">A propriedade `location` retorna um objeto `Location` que fornece os métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-437">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-437">Type</span></span>

*   <span data-ttu-id="43c0f-438">Cadeia de caracteres | [Localização](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="43c0f-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-439">Requirements</span></span>

|<span data-ttu-id="43c0f-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-440">Requirement</span></span>| <span data-ttu-id="43c0f-441">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-443">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-443">1.0</span></span>|
|[<span data-ttu-id="43c0f-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-445">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-446">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-447">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="43c0f-448">normalizedSubject :Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c0f-448">normalizedSubject :String</span></span>

<span data-ttu-id="43c0f-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="43c0f-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="43c0f-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-453">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-453">Type</span></span>

*   <span data-ttu-id="43c0f-454">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-455">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-455">Requirements</span></span>

|<span data-ttu-id="43c0f-456">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-456">Requirement</span></span>| <span data-ttu-id="43c0f-457">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-458">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-459">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-459">1.0</span></span>|
|[<span data-ttu-id="43c0f-460">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-461">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-462">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-463">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-464">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="43c0f-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="43c0f-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="43c0f-466">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-467">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-467">Type</span></span>

*   [<span data-ttu-id="43c0f-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="43c0f-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="43c0f-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-469">Requirements</span></span>

|<span data-ttu-id="43c0f-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-470">Requirement</span></span>| <span data-ttu-id="43c0f-471">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-472">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-473">1.3</span><span class="sxs-lookup"><span data-stu-id="43c0f-473">1.3</span></span>|
|[<span data-ttu-id="43c0f-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-475">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-476">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-477">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="43c0f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="43c0f-480">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="43c0f-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="43c0f-481">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-482">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-482">Read mode</span></span>

<span data-ttu-id="43c0f-483">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="43c0f-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-484">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-484">Compose mode</span></span>

<span data-ttu-id="43c0f-485">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="43c0f-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-486">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-486">Type</span></span>

*   <span data-ttu-id="43c0f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-488">Requirements</span></span>

|<span data-ttu-id="43c0f-489">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-489">Requirement</span></span>| <span data-ttu-id="43c0f-490">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-491">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-492">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-492">1.0</span></span>|
|[<span data-ttu-id="43c0f-493">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-494">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-495">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-496">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="43c0f-497">organizador:[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="43c0f-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="43c0f-p124">Obtém o endereço de email do organizador da reunião de uma reunião especificada. Modo somente leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-500">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-500">Type</span></span>

*   [<span data-ttu-id="43c0f-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43c0f-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="43c0f-502">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-502">Requirements</span></span>

|<span data-ttu-id="43c0f-503">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-503">Requirement</span></span>| <span data-ttu-id="43c0f-504">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-505">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-506">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-506">1.0</span></span>|
|[<span data-ttu-id="43c0f-507">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-508">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-509">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-510">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-511">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="43c0f-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="43c0f-513">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="43c0f-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="43c0f-514">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-515">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-515">Read mode</span></span>

<span data-ttu-id="43c0f-516">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="43c0f-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-517">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-517">Compose mode</span></span>

<span data-ttu-id="43c0f-518">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="43c0f-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-519">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-519">Type</span></span>

*   <span data-ttu-id="43c0f-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-521">Requirements</span></span>

|<span data-ttu-id="43c0f-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-522">Requirement</span></span>| <span data-ttu-id="43c0f-523">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-525">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-525">1.0</span></span>|
|[<span data-ttu-id="43c0f-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-527">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-528">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-529">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="43c0f-530">remetente :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="43c0f-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="43c0f-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="43c0f-p127">As propriedades [`from`](#from-emailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-535">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="43c0f-536">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-536">Type</span></span>

*   [<span data-ttu-id="43c0f-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="43c0f-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="43c0f-538">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-538">Requirements</span></span>

|<span data-ttu-id="43c0f-539">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-539">Requirement</span></span>| <span data-ttu-id="43c0f-540">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-541">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-542">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-542">1.0</span></span>|
|[<span data-ttu-id="43c0f-543">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-543">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-544">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-545">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-545">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-546">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-547">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="43c0f-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="43c0f-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="43c0f-549">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="43c0f-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="43c0f-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-552">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-552">Read mode</span></span>

<span data-ttu-id="43c0f-553">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-554">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-554">Compose mode</span></span>

<span data-ttu-id="43c0f-555">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="43c0f-556">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="43c0f-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="43c0f-557">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="43c0f-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-558">Type</span></span>

*   <span data-ttu-id="43c0f-559">Data | [Hora](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="43c0f-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-560">Requirements</span></span>

|<span data-ttu-id="43c0f-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-561">Requirement</span></span>| <span data-ttu-id="43c0f-562">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-564">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-564">1.0</span></span>|
|[<span data-ttu-id="43c0f-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-566">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-568">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="43c0f-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="43c0f-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="43c0f-570">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="43c0f-571">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="43c0f-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-572">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-572">Read mode</span></span>

<span data-ttu-id="43c0f-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-575">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-575">Compose mode</span></span>

<span data-ttu-id="43c0f-576">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="43c0f-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-577">Type</span></span>

*   <span data-ttu-id="43c0f-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="43c0f-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-579">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-579">Requirements</span></span>

|<span data-ttu-id="43c0f-580">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-580">Requirement</span></span>| <span data-ttu-id="43c0f-581">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-582">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-583">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-583">1.0</span></span>|
|[<span data-ttu-id="43c0f-584">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-585">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-586">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-587">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-587">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="43c0f-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="43c0f-589">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="43c0f-590">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="43c0f-591">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="43c0f-591">Read mode</span></span>

<span data-ttu-id="43c0f-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="43c0f-594">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="43c0f-594">Compose mode</span></span>

<span data-ttu-id="43c0f-595">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="43c0f-596">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-596">Type</span></span>

*   <span data-ttu-id="43c0f-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="43c0f-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-598">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-598">Requirements</span></span>

|<span data-ttu-id="43c0f-599">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-599">Requirement</span></span>| <span data-ttu-id="43c0f-600">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-601">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-602">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-602">1.0</span></span>|
|[<span data-ttu-id="43c0f-603">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-604">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-605">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-606">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="43c0f-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="43c0f-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="43c0f-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43c0f-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="43c0f-609">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="43c0f-610">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="43c0f-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="43c0f-611">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="43c0f-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-612">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-612">Parameters</span></span>

|<span data-ttu-id="43c0f-613">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-613">Name</span></span>| <span data-ttu-id="43c0f-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-614">Type</span></span>| <span data-ttu-id="43c0f-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-615">Attributes</span></span>| <span data-ttu-id="43c0f-616">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="43c0f-617">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-617">String</span></span>||<span data-ttu-id="43c0f-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="43c0f-620">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-620">String</span></span>||<span data-ttu-id="43c0f-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="43c0f-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-623">Object</span></span>| <span data-ttu-id="43c0f-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-624">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="43c0f-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-626">Object</span></span> | <span data-ttu-id="43c0f-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-627">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="43c0f-629">Booliano</span><span class="sxs-lookup"><span data-stu-id="43c0f-629">Boolean</span></span> | <span data-ttu-id="43c0f-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-630">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-631">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="43c0f-632">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-632">function</span></span>| <span data-ttu-id="43c0f-633">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-633">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-634">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43c0f-635">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="43c0f-636">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="43c0f-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43c0f-637">Erros</span><span class="sxs-lookup"><span data-stu-id="43c0f-637">Errors</span></span>

| <span data-ttu-id="43c0f-638">Código de erro</span><span class="sxs-lookup"><span data-stu-id="43c0f-638">Error code</span></span> | <span data-ttu-id="43c0f-639">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="43c0f-640">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="43c0f-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="43c0f-641">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="43c0f-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="43c0f-642">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-643">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-643">Requirements</span></span>

|<span data-ttu-id="43c0f-644">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-644">Requirement</span></span>| <span data-ttu-id="43c0f-645">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-646">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-647">1.1</span><span class="sxs-lookup"><span data-stu-id="43c0f-647">1.1</span></span>|
|[<span data-ttu-id="43c0f-648">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-650">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-651">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="43c0f-652">Exemplos</span><span class="sxs-lookup"><span data-stu-id="43c0f-652">Examples</span></span>

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

<span data-ttu-id="43c0f-653">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="43c0f-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43c0f-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="43c0f-655">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="43c0f-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="43c0f-659">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="43c0f-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="43c0f-660">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-661">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-661">Parameters</span></span>

|<span data-ttu-id="43c0f-662">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-662">Name</span></span>| <span data-ttu-id="43c0f-663">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-663">Type</span></span>| <span data-ttu-id="43c0f-664">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-664">Attributes</span></span>| <span data-ttu-id="43c0f-665">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="43c0f-666">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-666">String</span></span>||<span data-ttu-id="43c0f-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="43c0f-669">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c0f-669">String</span></span>||<span data-ttu-id="43c0f-670">O assunto do item a ser anexado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-670">The subject of the item to be attached.</span></span> <span data-ttu-id="43c0f-671">O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="43c0f-672">Object</span><span class="sxs-lookup"><span data-stu-id="43c0f-672">Object</span></span>| <span data-ttu-id="43c0f-673">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-673">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-674">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43c0f-675">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-675">Object</span></span>| <span data-ttu-id="43c0f-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-676">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-677">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43c0f-678">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-678">function</span></span>| <span data-ttu-id="43c0f-679">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-679">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-680">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43c0f-681">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="43c0f-682">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="43c0f-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43c0f-683">Erros</span><span class="sxs-lookup"><span data-stu-id="43c0f-683">Errors</span></span>

| <span data-ttu-id="43c0f-684">Código de erro</span><span class="sxs-lookup"><span data-stu-id="43c0f-684">Error code</span></span> | <span data-ttu-id="43c0f-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="43c0f-686">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-687">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-687">Requirements</span></span>

|<span data-ttu-id="43c0f-688">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-688">Requirement</span></span>| <span data-ttu-id="43c0f-689">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-690">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-691">1.1</span><span class="sxs-lookup"><span data-stu-id="43c0f-691">1.1</span></span>|
|[<span data-ttu-id="43c0f-692">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-694">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-695">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-696">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-696">Example</span></span>

<span data-ttu-id="43c0f-697">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="43c0f-698">close()</span><span class="sxs-lookup"><span data-stu-id="43c0f-698">close()</span></span>

<span data-ttu-id="43c0f-699">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="43c0f-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="43c0f-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-702">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="43c0f-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="43c0f-703">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="43c0f-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-704">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-704">Requirements</span></span>

|<span data-ttu-id="43c0f-705">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-705">Requirement</span></span>| <span data-ttu-id="43c0f-706">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-707">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-708">1.3</span><span class="sxs-lookup"><span data-stu-id="43c0f-708">1.3</span></span>|
|[<span data-ttu-id="43c0f-709">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-710">Restrito</span><span class="sxs-lookup"><span data-stu-id="43c0f-710">Restricted</span></span>|
|[<span data-ttu-id="43c0f-711">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-712">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="43c0f-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="43c0f-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="43c0f-714">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-715">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43c0f-716">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="43c0f-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="43c0f-717">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="43c0f-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="43c0f-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-721">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-721">Parameters</span></span>

| <span data-ttu-id="43c0f-722">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-722">Name</span></span> | <span data-ttu-id="43c0f-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-723">Type</span></span> | <span data-ttu-id="43c0f-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-724">Attributes</span></span> | <span data-ttu-id="43c0f-725">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="43c0f-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="43c0f-726">String &#124; Object</span></span>| |<span data-ttu-id="43c0f-727">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="43c0f-727">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="43c0f-728">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="43c0f-728">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="43c0f-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="43c0f-729">**OR**</span></span><br/><span data-ttu-id="43c0f-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="43c0f-732">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-732">String</span></span> | <span data-ttu-id="43c0f-733">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-733">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="43c0f-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43c0f-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-737">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-738">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="43c0f-739">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-739">String</span></span> | | <span data-ttu-id="43c0f-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="43c0f-742">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-742">String</span></span> | | <span data-ttu-id="43c0f-743">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="43c0f-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="43c0f-744">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-744">String</span></span> | | <span data-ttu-id="43c0f-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="43c0f-747">Booliano</span><span class="sxs-lookup"><span data-stu-id="43c0f-747">Boolean</span></span> | | <span data-ttu-id="43c0f-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="43c0f-750">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-750">String</span></span> | | <span data-ttu-id="43c0f-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="43c0f-754">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-754">function</span></span> | <span data-ttu-id="43c0f-755">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-755">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-756">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-757">Requirements</span></span>

|<span data-ttu-id="43c0f-758">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-758">Requirement</span></span>| <span data-ttu-id="43c0f-759">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-760">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-761">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-761">1.0</span></span>|
|[<span data-ttu-id="43c0f-762">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-763">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-764">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-765">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="43c0f-766">Exemplos</span><span class="sxs-lookup"><span data-stu-id="43c0f-766">Examples</span></span>

<span data-ttu-id="43c0f-767">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="43c0f-768">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="43c0f-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="43c0f-769">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="43c0f-770">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="43c0f-771">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="43c0f-772">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="43c0f-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="43c0f-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="43c0f-774">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-775">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43c0f-776">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="43c0f-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="43c0f-777">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="43c0f-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="43c0f-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-781">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-781">Parameters</span></span>

| <span data-ttu-id="43c0f-782">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-782">Name</span></span> | <span data-ttu-id="43c0f-783">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-783">Type</span></span> | <span data-ttu-id="43c0f-784">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-784">Attributes</span></span> | <span data-ttu-id="43c0f-785">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="43c0f-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="43c0f-786">String &#124; Object</span></span>| | <span data-ttu-id="43c0f-787">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta.</span><span class="sxs-lookup"><span data-stu-id="43c0f-787">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="43c0f-788">A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="43c0f-788">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="43c0f-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="43c0f-789">**OR**</span></span><br/><span data-ttu-id="43c0f-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="43c0f-792">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-792">String</span></span> | <span data-ttu-id="43c0f-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-793">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="43c0f-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43c0f-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-797">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-798">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="43c0f-799">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-799">String</span></span> | | <span data-ttu-id="43c0f-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="43c0f-802">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-802">String</span></span> | | <span data-ttu-id="43c0f-803">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="43c0f-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="43c0f-804">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-804">String</span></span> | | <span data-ttu-id="43c0f-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="43c0f-807">Booliano</span><span class="sxs-lookup"><span data-stu-id="43c0f-807">Boolean</span></span> | | <span data-ttu-id="43c0f-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="43c0f-810">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-810">String</span></span> | | <span data-ttu-id="43c0f-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="43c0f-814">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-814">function</span></span> | <span data-ttu-id="43c0f-815">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-815">&lt;optional&gt;</span></span> | <span data-ttu-id="43c0f-816">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-817">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-817">Requirements</span></span>

|<span data-ttu-id="43c0f-818">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-818">Requirement</span></span>| <span data-ttu-id="43c0f-819">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-820">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-821">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-821">1.0</span></span>|
|[<span data-ttu-id="43c0f-822">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-823">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-824">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-825">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="43c0f-826">Exemplos</span><span class="sxs-lookup"><span data-stu-id="43c0f-826">Examples</span></span>

<span data-ttu-id="43c0f-827">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="43c0f-828">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="43c0f-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="43c0f-829">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="43c0f-830">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="43c0f-831">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="43c0f-832">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="43c0f-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="43c0f-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="43c0f-834">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-835">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-836">Requirements</span></span>

|<span data-ttu-id="43c0f-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-837">Requirement</span></span>| <span data-ttu-id="43c0f-838">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-840">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-840">1.0</span></span>|
|[<span data-ttu-id="43c0f-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-842">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-844">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-845">Returns:</span></span>

<span data-ttu-id="43c0f-846">Tipo: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="43c0f-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="43c0f-847">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-847">Example</span></span>

<span data-ttu-id="43c0f-848">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="43c0f-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="43c0f-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="43c0f-850">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-851">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-852">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-852">Parameters</span></span>

|<span data-ttu-id="43c0f-853">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-853">Name</span></span>| <span data-ttu-id="43c0f-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-854">Type</span></span>| <span data-ttu-id="43c0f-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="43c0f-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="43c0f-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="43c0f-857">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="43c0f-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-858">Requirements</span></span>

|<span data-ttu-id="43c0f-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-859">Requirement</span></span>| <span data-ttu-id="43c0f-860">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-862">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-862">1.0</span></span>|
|[<span data-ttu-id="43c0f-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-864">Restrito</span><span class="sxs-lookup"><span data-stu-id="43c0f-864">Restricted</span></span>|
|[<span data-ttu-id="43c0f-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-866">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-867">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-867">Returns:</span></span>

<span data-ttu-id="43c0f-868">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="43c0f-869">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="43c0f-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="43c0f-870">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="43c0f-871">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="43c0f-872">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="43c0f-872">Value of `entityType`</span></span> | <span data-ttu-id="43c0f-873">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="43c0f-873">Type of objects in returned array</span></span> | <span data-ttu-id="43c0f-874">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="43c0f-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="43c0f-875">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-875">String</span></span> | <span data-ttu-id="43c0f-876">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="43c0f-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="43c0f-877">Contato</span><span class="sxs-lookup"><span data-stu-id="43c0f-877">Contact</span></span> | <span data-ttu-id="43c0f-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43c0f-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="43c0f-879">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-879">String</span></span> | <span data-ttu-id="43c0f-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43c0f-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="43c0f-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="43c0f-881">MeetingSuggestion</span></span> | <span data-ttu-id="43c0f-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43c0f-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="43c0f-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="43c0f-883">PhoneNumber</span></span> | <span data-ttu-id="43c0f-884">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="43c0f-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="43c0f-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="43c0f-885">TaskSuggestion</span></span> | <span data-ttu-id="43c0f-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="43c0f-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="43c0f-887">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-887">String</span></span> | <span data-ttu-id="43c0f-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="43c0f-888">**Restricted**</span></span> |

<span data-ttu-id="43c0f-889">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="43c0f-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="43c0f-890">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-890">Example</span></span>

<span data-ttu-id="43c0f-891">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="43c0f-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="43c0f-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="43c0f-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="43c0f-893">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-894">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43c0f-895">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-896">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-896">Parameters</span></span>

|<span data-ttu-id="43c0f-897">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-897">Name</span></span>| <span data-ttu-id="43c0f-898">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-898">Type</span></span>| <span data-ttu-id="43c0f-899">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="43c0f-900">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-900">String</span></span>|<span data-ttu-id="43c0f-901">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="43c0f-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-902">Requirements</span></span>

|<span data-ttu-id="43c0f-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-903">Requirement</span></span>| <span data-ttu-id="43c0f-904">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-906">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-906">1.0</span></span>|
|[<span data-ttu-id="43c0f-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-908">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-910">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-911">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-911">Returns:</span></span>

<span data-ttu-id="43c0f-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="43c0f-914">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="43c0f-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="43c0f-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="43c0f-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="43c0f-916">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-917">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43c0f-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="43c0f-921">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="43c0f-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="43c0f-922">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="43c0f-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c0f-926">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-926">Requirements</span></span>

|<span data-ttu-id="43c0f-927">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-927">Requirement</span></span>| <span data-ttu-id="43c0f-928">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-929">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-930">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-930">1.0</span></span>|
|[<span data-ttu-id="43c0f-931">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-932">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-933">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-934">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-935">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-935">Returns:</span></span>

<span data-ttu-id="43c0f-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="43c0f-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43c0f-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43c0f-939">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43c0f-940">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-940">Example</span></span>

<span data-ttu-id="43c0f-941">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="43c0f-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="43c0f-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="43c0f-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="43c0f-943">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-944">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="43c0f-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43c0f-945">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="43c0f-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-948">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-948">Parameters</span></span>

|<span data-ttu-id="43c0f-949">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-949">Name</span></span>| <span data-ttu-id="43c0f-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-950">Type</span></span>| <span data-ttu-id="43c0f-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="43c0f-952">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-952">String</span></span>|<span data-ttu-id="43c0f-953">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="43c0f-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-954">Requirements</span></span>

|<span data-ttu-id="43c0f-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-955">Requirement</span></span>| <span data-ttu-id="43c0f-956">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-957">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-958">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-958">1.0</span></span>|
|[<span data-ttu-id="43c0f-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-960">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-962">Read</span><span class="sxs-lookup"><span data-stu-id="43c0f-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-963">Returns:</span></span>

<span data-ttu-id="43c0f-964">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="43c0f-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43c0f-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43c0f-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="43c0f-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43c0f-967">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="43c0f-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="43c0f-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="43c0f-969">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="43c0f-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-972">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-972">Parameters</span></span>

|<span data-ttu-id="43c0f-973">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-973">Name</span></span>| <span data-ttu-id="43c0f-974">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-974">Type</span></span>| <span data-ttu-id="43c0f-975">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-975">Attributes</span></span>| <span data-ttu-id="43c0f-976">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="43c0f-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="43c0f-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="43c0f-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="43c0f-981">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-981">Object</span></span>| <span data-ttu-id="43c0f-982">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-982">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-983">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43c0f-984">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-984">Object</span></span>| <span data-ttu-id="43c0f-985">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-985">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-986">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43c0f-987">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-987">function</span></span>||<span data-ttu-id="43c0f-988">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43c0f-989">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="43c0f-990">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-991">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-991">Requirements</span></span>

|<span data-ttu-id="43c0f-992">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-992">Requirement</span></span>| <span data-ttu-id="43c0f-993">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-994">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-995">1.2</span><span class="sxs-lookup"><span data-stu-id="43c0f-995">1.2</span></span>|
|[<span data-ttu-id="43c0f-996">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-998">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-999">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="43c0f-1000">Retorna:</span><span class="sxs-lookup"><span data-stu-id="43c0f-1000">Returns:</span></span>

<span data-ttu-id="43c0f-1001">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="43c0f-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="43c0f-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43c0f-1003">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="43c0f-1004">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="43c0f-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43c0f-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="43c0f-1006">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="43c0f-p163">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-1010">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-1010">Parameters</span></span>

|<span data-ttu-id="43c0f-1011">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-1011">Name</span></span>| <span data-ttu-id="43c0f-1012">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1012">Type</span></span>| <span data-ttu-id="43c0f-1013">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1013">Attributes</span></span>| <span data-ttu-id="43c0f-1014">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43c0f-1015">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-1015">function</span></span>||<span data-ttu-id="43c0f-1016">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43c0f-1017">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="43c0f-1018">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="43c0f-1019">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1019">Object</span></span>| <span data-ttu-id="43c0f-1020">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1021">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="43c0f-1022">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1023">Requirements</span></span>

|<span data-ttu-id="43c0f-1024">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-1024">Requirement</span></span>| <span data-ttu-id="43c0f-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-1026">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="43c0f-1027">1.0</span></span>|
|[<span data-ttu-id="43c0f-1028">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-1029">ReadItem</span></span>|
|[<span data-ttu-id="43c0f-1030">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="43c0f-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-1031">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c0f-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1032">Example</span></span>

<span data-ttu-id="43c0f-p166">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="43c0f-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43c0f-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="43c0f-1037">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="43c0f-p167">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-1042">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-1042">Parameters</span></span>

|<span data-ttu-id="43c0f-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-1043">Name</span></span>| <span data-ttu-id="43c0f-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1044">Type</span></span>| <span data-ttu-id="43c0f-1045">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1045">Attributes</span></span>| <span data-ttu-id="43c0f-1046">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="43c0f-1047">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-1047">String</span></span>||<span data-ttu-id="43c0f-1048">O identificador do anexo a remover.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="43c0f-1049">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1049">Object</span></span>| <span data-ttu-id="43c0f-1050">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1051">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43c0f-1052">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1052">Object</span></span>| <span data-ttu-id="43c0f-1053">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1054">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43c0f-1055">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-1055">function</span></span>| <span data-ttu-id="43c0f-1056">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1057">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="43c0f-1058">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43c0f-1059">Erros</span><span class="sxs-lookup"><span data-stu-id="43c0f-1059">Errors</span></span>

| <span data-ttu-id="43c0f-1060">Código de erro</span><span class="sxs-lookup"><span data-stu-id="43c0f-1060">Error code</span></span> | <span data-ttu-id="43c0f-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="43c0f-1062">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-1063">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1063">Requirements</span></span>

|<span data-ttu-id="43c0f-1064">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-1064">Requirement</span></span>| <span data-ttu-id="43c0f-1065">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-1066">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="43c0f-1067">1.1</span></span>|
|[<span data-ttu-id="43c0f-1068">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1068">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-1070">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-1070">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-1071">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-1072">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1072">Example</span></span>

<span data-ttu-id="43c0f-1073">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="43c0f-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="43c0f-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="43c0f-1075">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="43c0f-p168">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-1079">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="43c0f-1080">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="43c0f-p170">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="43c0f-1084">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="43c0f-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="43c0f-1085">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1085">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="43c0f-1086">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1086">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="43c0f-1087">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1087">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-1088">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-1088">Parameters</span></span>

|<span data-ttu-id="43c0f-1089">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-1089">Name</span></span>| <span data-ttu-id="43c0f-1090">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1090">Type</span></span>| <span data-ttu-id="43c0f-1091">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1091">Attributes</span></span>| <span data-ttu-id="43c0f-1092">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-1092">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="43c0f-1093">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1093">Object</span></span>| <span data-ttu-id="43c0f-1094">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1095">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43c0f-1096">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1096">Object</span></span>| <span data-ttu-id="43c0f-1097">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1098">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="43c0f-1099">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-1099">function</span></span>||<span data-ttu-id="43c0f-1100">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-1100">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43c0f-1101">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1101">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c0f-1102">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1102">Requirements</span></span>

|<span data-ttu-id="43c0f-1103">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-1103">Requirement</span></span>| <span data-ttu-id="43c0f-1104">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-1104">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-1105">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-1105">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-1106">1.3</span><span class="sxs-lookup"><span data-stu-id="43c0f-1106">1.3</span></span>|
|[<span data-ttu-id="43c0f-1107">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1107">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-1108">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-1108">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-1109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-1109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-1110">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-1110">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="43c0f-1111">Exemplos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1111">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="43c0f-p172">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="43c0f-1114">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="43c0f-1114">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="43c0f-1115">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1115">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="43c0f-p173">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43c0f-1119">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="43c0f-1119">Parameters</span></span>

|<span data-ttu-id="43c0f-1120">Nome</span><span class="sxs-lookup"><span data-stu-id="43c0f-1120">Name</span></span>| <span data-ttu-id="43c0f-1121">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1121">Type</span></span>| <span data-ttu-id="43c0f-1122">Atributos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1122">Attributes</span></span>| <span data-ttu-id="43c0f-1123">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c0f-1123">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="43c0f-1124">String</span><span class="sxs-lookup"><span data-stu-id="43c0f-1124">String</span></span>||<span data-ttu-id="43c0f-p174">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="43c0f-1128">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1128">Object</span></span>| <span data-ttu-id="43c0f-1129">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1129">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1130">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1130">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="43c0f-1131">Objeto</span><span class="sxs-lookup"><span data-stu-id="43c0f-1131">Object</span></span>| <span data-ttu-id="43c0f-1132">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1132">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-1133">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1133">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="43c0f-1134">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="43c0f-1134">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="43c0f-1135">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="43c0f-1135">&lt;optional&gt;</span></span>|<span data-ttu-id="43c0f-p175">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="43c0f-p176">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="43c0f-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="43c0f-1140">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="43c0f-1140">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="43c0f-1141">function</span><span class="sxs-lookup"><span data-stu-id="43c0f-1141">function</span></span>||<span data-ttu-id="43c0f-1142">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="43c0f-1142">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43c0f-1143">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c0f-1143">Requirements</span></span>

|<span data-ttu-id="43c0f-1144">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c0f-1144">Requirement</span></span>| <span data-ttu-id="43c0f-1145">Valor</span><span class="sxs-lookup"><span data-stu-id="43c0f-1145">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c0f-1146">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c0f-1146">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43c0f-1147">1.2</span><span class="sxs-lookup"><span data-stu-id="43c0f-1147">1.2</span></span>|
|[<span data-ttu-id="43c0f-1148">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1148">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43c0f-1149">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="43c0f-1149">ReadWriteItem</span></span>|
|[<span data-ttu-id="43c0f-1150">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c0f-1150">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43c0f-1151">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c0f-1151">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="43c0f-1152">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43c0f-1152">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
