
# <a name="item"></a><span data-ttu-id="07351-101">item</span><span class="sxs-lookup"><span data-stu-id="07351-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="07351-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="07351-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="07351-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="07351-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-105">Requirements</span></span>

|<span data-ttu-id="07351-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-106">Requirement</span></span>| <span data-ttu-id="07351-107">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-109">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-109">1.0</span></span>|
|[<span data-ttu-id="07351-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="07351-111">Restricted</span></span>|
|[<span data-ttu-id="07351-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="07351-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="07351-114">Members and methods</span></span>

| <span data-ttu-id="07351-115">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-115">Member</span></span> | <span data-ttu-id="07351-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="07351-117">attachments</span><span class="sxs-lookup"><span data-stu-id="07351-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="07351-118">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-118">Member</span></span> |
| [<span data-ttu-id="07351-119">bcc</span><span class="sxs-lookup"><span data-stu-id="07351-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="07351-120">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-120">Member</span></span> |
| [<span data-ttu-id="07351-121">body</span><span class="sxs-lookup"><span data-stu-id="07351-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="07351-122">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-122">Member</span></span> |
| [<span data-ttu-id="07351-123">cc</span><span class="sxs-lookup"><span data-stu-id="07351-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="07351-124">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-124">Member</span></span> |
| [<span data-ttu-id="07351-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="07351-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="07351-126">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-126">Member</span></span> |
| [<span data-ttu-id="07351-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="07351-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="07351-128">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-128">Member</span></span> |
| [<span data-ttu-id="07351-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="07351-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="07351-130">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-130">Member</span></span> |
| [<span data-ttu-id="07351-131">end</span><span class="sxs-lookup"><span data-stu-id="07351-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="07351-132">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-132">Member</span></span> |
| [<span data-ttu-id="07351-133">from</span><span class="sxs-lookup"><span data-stu-id="07351-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="07351-134">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-134">Member</span></span> |
| [<span data-ttu-id="07351-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="07351-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="07351-136">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-136">Member</span></span> |
| [<span data-ttu-id="07351-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="07351-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="07351-138">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-138">Member</span></span> |
| [<span data-ttu-id="07351-139">itemId</span><span class="sxs-lookup"><span data-stu-id="07351-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="07351-140">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-140">Member</span></span> |
| [<span data-ttu-id="07351-141">itemType</span><span class="sxs-lookup"><span data-stu-id="07351-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="07351-142">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-142">Member</span></span> |
| [<span data-ttu-id="07351-143">location</span><span class="sxs-lookup"><span data-stu-id="07351-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="07351-144">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-144">Member</span></span> |
| [<span data-ttu-id="07351-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="07351-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="07351-146">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-146">Member</span></span> |
| [<span data-ttu-id="07351-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="07351-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="07351-148">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-148">Member</span></span> |
| [<span data-ttu-id="07351-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="07351-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="07351-150">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-150">Member</span></span> |
| [<span data-ttu-id="07351-151">organizer</span><span class="sxs-lookup"><span data-stu-id="07351-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="07351-152">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-152">Member</span></span> |
| [<span data-ttu-id="07351-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="07351-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="07351-154">Member</span><span class="sxs-lookup"><span data-stu-id="07351-154">Member</span></span> |
| [<span data-ttu-id="07351-155">sender</span><span class="sxs-lookup"><span data-stu-id="07351-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="07351-156">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-156">Member</span></span> |
| [<span data-ttu-id="07351-157">start</span><span class="sxs-lookup"><span data-stu-id="07351-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="07351-158">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-158">Member</span></span> |
| [<span data-ttu-id="07351-159">subject</span><span class="sxs-lookup"><span data-stu-id="07351-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="07351-160">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-160">Member</span></span> |
| [<span data-ttu-id="07351-161">to</span><span class="sxs-lookup"><span data-stu-id="07351-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="07351-162">Membro</span><span class="sxs-lookup"><span data-stu-id="07351-162">Member</span></span> |
| [<span data-ttu-id="07351-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07351-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="07351-164">Método</span><span class="sxs-lookup"><span data-stu-id="07351-164">Method</span></span> |
| [<span data-ttu-id="07351-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07351-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="07351-166">Método</span><span class="sxs-lookup"><span data-stu-id="07351-166">Method</span></span> |
| [<span data-ttu-id="07351-167">close</span><span class="sxs-lookup"><span data-stu-id="07351-167">close</span></span>](#close) | <span data-ttu-id="07351-168">Método</span><span class="sxs-lookup"><span data-stu-id="07351-168">Method</span></span> |
| [<span data-ttu-id="07351-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="07351-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="07351-170">Método</span><span class="sxs-lookup"><span data-stu-id="07351-170">Method</span></span> |
| [<span data-ttu-id="07351-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="07351-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="07351-172">Método</span><span class="sxs-lookup"><span data-stu-id="07351-172">Method</span></span> |
| [<span data-ttu-id="07351-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="07351-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="07351-174">Método</span><span class="sxs-lookup"><span data-stu-id="07351-174">Method</span></span> |
| [<span data-ttu-id="07351-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="07351-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="07351-176">Método</span><span class="sxs-lookup"><span data-stu-id="07351-176">Method</span></span> |
| [<span data-ttu-id="07351-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="07351-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="07351-178">Método</span><span class="sxs-lookup"><span data-stu-id="07351-178">Method</span></span> |
| [<span data-ttu-id="07351-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="07351-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="07351-180">Método</span><span class="sxs-lookup"><span data-stu-id="07351-180">Method</span></span> |
| [<span data-ttu-id="07351-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="07351-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="07351-182">Método</span><span class="sxs-lookup"><span data-stu-id="07351-182">Method</span></span> |
| [<span data-ttu-id="07351-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07351-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="07351-184">Método</span><span class="sxs-lookup"><span data-stu-id="07351-184">Method</span></span> |
| [<span data-ttu-id="07351-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="07351-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="07351-186">Método</span><span class="sxs-lookup"><span data-stu-id="07351-186">Method</span></span> |
| [<span data-ttu-id="07351-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07351-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="07351-188">Método</span><span class="sxs-lookup"><span data-stu-id="07351-188">Method</span></span> |
| [<span data-ttu-id="07351-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="07351-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="07351-190">Método</span><span class="sxs-lookup"><span data-stu-id="07351-190">Method</span></span> |
| [<span data-ttu-id="07351-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07351-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="07351-192">Método</span><span class="sxs-lookup"><span data-stu-id="07351-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="07351-193">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-193">Example</span></span>

<span data-ttu-id="07351-194">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="07351-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="07351-195">Membros</span><span class="sxs-lookup"><span data-stu-id="07351-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="07351-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07351-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="07351-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-199">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="07351-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="07351-200">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="07351-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="07351-201">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-201">Type:</span></span>

*   <span data-ttu-id="07351-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07351-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-203">Requirements</span></span>

|<span data-ttu-id="07351-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-204">Requirement</span></span>| <span data-ttu-id="07351-205">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-207">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-207">1.0</span></span>|
|[<span data-ttu-id="07351-208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-209">ReadItem</span></span>|
|[<span data-ttu-id="07351-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-211">Read</span><span class="sxs-lookup"><span data-stu-id="07351-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-212">Example</span></span>

<span data-ttu-id="07351-213">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="07351-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="07351-215">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="07351-216">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="07351-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-217">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-217">Type:</span></span>

*   [<span data-ttu-id="07351-218">Destinatários</span><span class="sxs-lookup"><span data-stu-id="07351-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="07351-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-219">Requirements</span></span>

|<span data-ttu-id="07351-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-220">Requirement</span></span>| <span data-ttu-id="07351-221">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-223">1.1</span><span class="sxs-lookup"><span data-stu-id="07351-223">1.1</span></span>|
|[<span data-ttu-id="07351-224">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-225">ReadItem</span></span>|
|[<span data-ttu-id="07351-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-227">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-228">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-228">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="07351-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="07351-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="07351-230">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="07351-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-231">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-231">Type:</span></span>

*   [<span data-ttu-id="07351-232">Corpo</span><span class="sxs-lookup"><span data-stu-id="07351-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="07351-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-233">Requirements</span></span>

|<span data-ttu-id="07351-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-234">Requirement</span></span>| <span data-ttu-id="07351-235">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-237">1.1</span><span class="sxs-lookup"><span data-stu-id="07351-237">1.1</span></span>|
|[<span data-ttu-id="07351-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-239">ReadItem</span></span>|
|[<span data-ttu-id="07351-240">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-241">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="07351-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="07351-243">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="07351-244">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-245">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-245">Read mode</span></span>

<span data-ttu-id="07351-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="07351-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-248">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-248">Compose mode</span></span>

<span data-ttu-id="07351-249">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-250">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-250">Type:</span></span>

*   <span data-ttu-id="07351-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-252">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-252">Requirements</span></span>

|<span data-ttu-id="07351-253">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-253">Requirement</span></span>| <span data-ttu-id="07351-254">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-255">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-256">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-256">1.0</span></span>|
|[<span data-ttu-id="07351-257">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-258">ReadItem</span></span>|
|[<span data-ttu-id="07351-259">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-260">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-261">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-261">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="07351-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="07351-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="07351-263">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="07351-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="07351-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="07351-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="07351-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="07351-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-268">Type:</span></span>

*   <span data-ttu-id="07351-269">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07351-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-270">Requirements</span></span>

|<span data-ttu-id="07351-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-271">Requirement</span></span>| <span data-ttu-id="07351-272">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-274">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-274">1.0</span></span>|
|[<span data-ttu-id="07351-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-276">ReadItem</span></span>|
|[<span data-ttu-id="07351-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-278">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="07351-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="07351-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="07351-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-282">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-282">Type:</span></span>

*   <span data-ttu-id="07351-283">Data</span><span class="sxs-lookup"><span data-stu-id="07351-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-284">Requirements</span></span>

|<span data-ttu-id="07351-285">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-285">Requirement</span></span>| <span data-ttu-id="07351-286">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-287">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-288">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-288">1.0</span></span>|
|[<span data-ttu-id="07351-289">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-290">ReadItem</span></span>|
|[<span data-ttu-id="07351-291">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-292">Read</span><span class="sxs-lookup"><span data-stu-id="07351-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-293">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-293">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="07351-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="07351-294">dateTimeModified :Date</span></span>

<span data-ttu-id="07351-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-297">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-298">Type:</span></span>

*   <span data-ttu-id="07351-299">Data</span><span class="sxs-lookup"><span data-stu-id="07351-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-300">Requirements</span></span>

|<span data-ttu-id="07351-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-301">Requirement</span></span>| <span data-ttu-id="07351-302">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-304">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-304">1.0</span></span>|
|[<span data-ttu-id="07351-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-306">ReadItem</span></span>|
|[<span data-ttu-id="07351-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-308">Read</span><span class="sxs-lookup"><span data-stu-id="07351-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-309">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="07351-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="07351-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="07351-311">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="07351-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="07351-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="07351-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-314">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-314">Read mode</span></span>

<span data-ttu-id="07351-315">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="07351-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-316">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-316">Compose mode</span></span>

<span data-ttu-id="07351-317">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="07351-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="07351-318">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="07351-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-319">Type:</span></span>

*   <span data-ttu-id="07351-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="07351-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-321">Requirements</span></span>

|<span data-ttu-id="07351-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-322">Requirement</span></span>| <span data-ttu-id="07351-323">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-325">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-325">1.0</span></span>|
|[<span data-ttu-id="07351-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-327">ReadItem</span></span>|
|[<span data-ttu-id="07351-328">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-329">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-330">Example</span></span>

<span data-ttu-id="07351-331">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="07351-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="07351-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="07351-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="07351-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="07351-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="07351-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-337">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="07351-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-338">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-338">Type:</span></span>

*   [<span data-ttu-id="07351-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07351-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="07351-340">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-340">Requirements</span></span>

|<span data-ttu-id="07351-341">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-341">Requirement</span></span>| <span data-ttu-id="07351-342">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-343">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-344">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-344">1.0</span></span>|
|[<span data-ttu-id="07351-345">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-346">ReadItem</span></span>|
|[<span data-ttu-id="07351-347">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-348">Read</span><span class="sxs-lookup"><span data-stu-id="07351-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="07351-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="07351-349">internetMessageId :String</span></span>

<span data-ttu-id="07351-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-352">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-352">Type:</span></span>

*   <span data-ttu-id="07351-353">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07351-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-354">Requirements</span></span>

|<span data-ttu-id="07351-355">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-355">Requirement</span></span>| <span data-ttu-id="07351-356">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-357">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-358">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-358">1.0</span></span>|
|[<span data-ttu-id="07351-359">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-360">ReadItem</span></span>|
|[<span data-ttu-id="07351-361">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-362">Read</span><span class="sxs-lookup"><span data-stu-id="07351-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-363">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-363">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="07351-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="07351-364">itemClass :String</span></span>

<span data-ttu-id="07351-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="07351-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="07351-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-369">Type</span></span> | <span data-ttu-id="07351-370">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-370">Description</span></span> | <span data-ttu-id="07351-371">classe de item</span><span class="sxs-lookup"><span data-stu-id="07351-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="07351-372">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="07351-372">Appointment items</span></span> | <span data-ttu-id="07351-373">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="07351-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="07351-374">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="07351-374">Message items</span></span> | <span data-ttu-id="07351-375">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="07351-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="07351-376">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="07351-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-377">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-377">Type:</span></span>

*   <span data-ttu-id="07351-378">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07351-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-379">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-379">Requirements</span></span>

|<span data-ttu-id="07351-380">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-380">Requirement</span></span>| <span data-ttu-id="07351-381">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-382">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-383">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-383">1.0</span></span>|
|[<span data-ttu-id="07351-384">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-385">ReadItem</span></span>|
|[<span data-ttu-id="07351-386">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-387">Read</span><span class="sxs-lookup"><span data-stu-id="07351-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-388">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-388">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="07351-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="07351-389">(nullable) itemId :String</span></span>

<span data-ttu-id="07351-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-392">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="07351-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="07351-393">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="07351-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="07351-394">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="07351-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="07351-395">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="07351-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="07351-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-398">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-398">Type:</span></span>

*   <span data-ttu-id="07351-399">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07351-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-400">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-400">Requirements</span></span>

|<span data-ttu-id="07351-401">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-401">Requirement</span></span>| <span data-ttu-id="07351-402">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-403">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-404">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-404">1.0</span></span>|
|[<span data-ttu-id="07351-405">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-406">ReadItem</span></span>|
|[<span data-ttu-id="07351-407">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-408">Read</span><span class="sxs-lookup"><span data-stu-id="07351-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-409">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-409">Example</span></span>

<span data-ttu-id="07351-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="07351-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="07351-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="07351-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="07351-413">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="07351-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="07351-414">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-415">Type:</span></span>

*   [<span data-ttu-id="07351-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="07351-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="07351-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-417">Requirements</span></span>

|<span data-ttu-id="07351-418">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-418">Requirement</span></span>| <span data-ttu-id="07351-419">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-420">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-421">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-421">1.0</span></span>|
|[<span data-ttu-id="07351-422">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-423">ReadItem</span></span>|
|[<span data-ttu-id="07351-424">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-425">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-426">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-426">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="07351-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="07351-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="07351-428">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-429">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-429">Read mode</span></span>

<span data-ttu-id="07351-430">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-431">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-431">Compose mode</span></span>

<span data-ttu-id="07351-432">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-433">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-433">Type:</span></span>

*   <span data-ttu-id="07351-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="07351-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-435">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-435">Requirements</span></span>

|<span data-ttu-id="07351-436">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-436">Requirement</span></span>| <span data-ttu-id="07351-437">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-438">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-439">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-439">1.0</span></span>|
|[<span data-ttu-id="07351-440">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-441">ReadItem</span></span>|
|[<span data-ttu-id="07351-442">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-443">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-444">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-444">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="07351-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="07351-445">normalizedSubject :String</span></span>

<span data-ttu-id="07351-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="07351-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="07351-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-450">Type:</span></span>

*   <span data-ttu-id="07351-451">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="07351-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-452">Requirements</span></span>

|<span data-ttu-id="07351-453">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-453">Requirement</span></span>| <span data-ttu-id="07351-454">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-455">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-456">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-456">1.0</span></span>|
|[<span data-ttu-id="07351-457">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-458">ReadItem</span></span>|
|[<span data-ttu-id="07351-459">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-460">Read</span><span class="sxs-lookup"><span data-stu-id="07351-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-461">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-461">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="07351-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="07351-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="07351-463">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="07351-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-464">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-464">Type:</span></span>

*   [<span data-ttu-id="07351-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="07351-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="07351-466">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-466">Requirements</span></span>

|<span data-ttu-id="07351-467">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-467">Requirement</span></span>| <span data-ttu-id="07351-468">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-469">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-470">1.3</span><span class="sxs-lookup"><span data-stu-id="07351-470">1.3</span></span>|
|[<span data-ttu-id="07351-471">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-472">ReadItem</span></span>|
|[<span data-ttu-id="07351-473">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-474">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="07351-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="07351-476">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="07351-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="07351-477">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-478">Read mode</span></span>

<span data-ttu-id="07351-479">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="07351-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-480">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-480">Compose mode</span></span>

<span data-ttu-id="07351-481">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="07351-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-482">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-482">Type:</span></span>

*   <span data-ttu-id="07351-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-484">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-484">Requirements</span></span>

|<span data-ttu-id="07351-485">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-485">Requirement</span></span>| <span data-ttu-id="07351-486">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-487">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-488">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-488">1.0</span></span>|
|[<span data-ttu-id="07351-489">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-490">ReadItem</span></span>|
|[<span data-ttu-id="07351-491">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-492">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-493">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-493">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="07351-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="07351-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="07351-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-497">Type:</span></span>

*   [<span data-ttu-id="07351-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07351-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="07351-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-499">Requirements</span></span>

|<span data-ttu-id="07351-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-500">Requirement</span></span>| <span data-ttu-id="07351-501">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-503">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-503">1.0</span></span>|
|[<span data-ttu-id="07351-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-505">ReadItem</span></span>|
|[<span data-ttu-id="07351-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-507">Read</span><span class="sxs-lookup"><span data-stu-id="07351-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-508">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-508">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="07351-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="07351-510">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="07351-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="07351-511">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-512">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-512">Read mode</span></span>

<span data-ttu-id="07351-513">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="07351-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-514">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-514">Compose mode</span></span>

<span data-ttu-id="07351-515">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="07351-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-516">Type:</span></span>

*   <span data-ttu-id="07351-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-518">Requirements</span></span>

|<span data-ttu-id="07351-519">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-519">Requirement</span></span>| <span data-ttu-id="07351-520">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-521">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-522">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-522">1.0</span></span>|
|[<span data-ttu-id="07351-523">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-524">ReadItem</span></span>|
|[<span data-ttu-id="07351-525">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-526">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-527">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-527">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="07351-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="07351-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="07351-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="07351-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="07351-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="07351-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-533">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="07351-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-534">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-534">Type:</span></span>

*   [<span data-ttu-id="07351-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07351-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="07351-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-536">Requirements</span></span>

|<span data-ttu-id="07351-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-537">Requirement</span></span>| <span data-ttu-id="07351-538">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-540">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-540">1.0</span></span>|
|[<span data-ttu-id="07351-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-542">ReadItem</span></span>|
|[<span data-ttu-id="07351-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-544">Read</span><span class="sxs-lookup"><span data-stu-id="07351-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-545">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="07351-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="07351-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="07351-547">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="07351-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="07351-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="07351-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-550">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-550">Read mode</span></span>

<span data-ttu-id="07351-551">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="07351-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-552">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-552">Compose mode</span></span>

<span data-ttu-id="07351-553">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="07351-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="07351-554">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="07351-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-555">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-555">Type:</span></span>

*   <span data-ttu-id="07351-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="07351-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-557">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-557">Requirements</span></span>

|<span data-ttu-id="07351-558">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-558">Requirement</span></span>| <span data-ttu-id="07351-559">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-560">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-561">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-561">1.0</span></span>|
|[<span data-ttu-id="07351-562">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-563">ReadItem</span></span>|
|[<span data-ttu-id="07351-564">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-565">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-566">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-566">Example</span></span>

<span data-ttu-id="07351-567">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="07351-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="07351-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="07351-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="07351-569">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="07351-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="07351-570">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="07351-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-571">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-571">Read mode</span></span>

<span data-ttu-id="07351-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="07351-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="07351-574">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-574">Compose mode</span></span>

<span data-ttu-id="07351-575">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="07351-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07351-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-576">Type:</span></span>

*   <span data-ttu-id="07351-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="07351-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-578">Requirements</span></span>

|<span data-ttu-id="07351-579">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-579">Requirement</span></span>| <span data-ttu-id="07351-580">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-581">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-582">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-582">1.0</span></span>|
|[<span data-ttu-id="07351-583">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-584">ReadItem</span></span>|
|[<span data-ttu-id="07351-585">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-586">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="07351-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="07351-588">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="07351-589">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07351-590">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="07351-590">Read mode</span></span>

<span data-ttu-id="07351-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="07351-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07351-593">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="07351-593">Compose mode</span></span>

<span data-ttu-id="07351-594">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="07351-595">Tipo:</span><span class="sxs-lookup"><span data-stu-id="07351-595">Type:</span></span>

*   <span data-ttu-id="07351-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07351-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-597">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-597">Requirements</span></span>

|<span data-ttu-id="07351-598">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-598">Requirement</span></span>| <span data-ttu-id="07351-599">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-600">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-601">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-601">1.0</span></span>|
|[<span data-ttu-id="07351-602">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-603">ReadItem</span></span>|
|[<span data-ttu-id="07351-604">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-605">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-606">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-606">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="07351-607">Métodos</span><span class="sxs-lookup"><span data-stu-id="07351-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="07351-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07351-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07351-609">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="07351-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="07351-610">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="07351-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="07351-611">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="07351-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-612">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-612">Parameters:</span></span>

|<span data-ttu-id="07351-613">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-613">Name</span></span>| <span data-ttu-id="07351-614">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-614">Type</span></span>| <span data-ttu-id="07351-615">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-615">Attributes</span></span>| <span data-ttu-id="07351-616">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="07351-617">String</span><span class="sxs-lookup"><span data-stu-id="07351-617">String</span></span>||<span data-ttu-id="07351-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="07351-620">String</span><span class="sxs-lookup"><span data-stu-id="07351-620">String</span></span>||<span data-ttu-id="07351-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="07351-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-623">Object</span></span>| <span data-ttu-id="07351-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-624">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="07351-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-626">Object</span></span> | <span data-ttu-id="07351-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-627">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="07351-629">Booliano</span><span class="sxs-lookup"><span data-stu-id="07351-629">Boolean</span></span> | <span data-ttu-id="07351-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-630">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-631">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="07351-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="07351-632">function</span><span class="sxs-lookup"><span data-stu-id="07351-632">function</span></span>| <span data-ttu-id="07351-633">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-633">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-634">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07351-635">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="07351-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07351-636">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="07351-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07351-637">Erros</span><span class="sxs-lookup"><span data-stu-id="07351-637">Errors</span></span>

| <span data-ttu-id="07351-638">Código de erro</span><span class="sxs-lookup"><span data-stu-id="07351-638">Error code</span></span> | <span data-ttu-id="07351-639">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="07351-640">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="07351-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="07351-641">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="07351-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="07351-642">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="07351-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-643">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-643">Requirements</span></span>

|<span data-ttu-id="07351-644">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-644">Requirement</span></span>| <span data-ttu-id="07351-645">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-646">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-647">1.1</span><span class="sxs-lookup"><span data-stu-id="07351-647">1.1</span></span>|
|[<span data-ttu-id="07351-648">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-650">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-651">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07351-652">Exemplos</span><span class="sxs-lookup"><span data-stu-id="07351-652">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="07351-653">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="07351-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07351-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07351-655">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="07351-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="07351-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="07351-659">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="07351-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="07351-660">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="07351-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-661">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-661">Parameters:</span></span>

|<span data-ttu-id="07351-662">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-662">Name</span></span>| <span data-ttu-id="07351-663">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-663">Type</span></span>| <span data-ttu-id="07351-664">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-664">Attributes</span></span>| <span data-ttu-id="07351-665">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="07351-666">String</span><span class="sxs-lookup"><span data-stu-id="07351-666">String</span></span>||<span data-ttu-id="07351-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="07351-669">String</span><span class="sxs-lookup"><span data-stu-id="07351-669">String</span></span>||<span data-ttu-id="07351-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="07351-672">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-672">Object</span></span>| <span data-ttu-id="07351-673">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-673">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-674">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07351-675">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-675">Object</span></span>| <span data-ttu-id="07351-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-676">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-677">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07351-678">function</span><span class="sxs-lookup"><span data-stu-id="07351-678">function</span></span>| <span data-ttu-id="07351-679">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-679">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-680">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07351-681">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="07351-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07351-682">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="07351-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07351-683">Erros</span><span class="sxs-lookup"><span data-stu-id="07351-683">Errors</span></span>

| <span data-ttu-id="07351-684">Código de erro</span><span class="sxs-lookup"><span data-stu-id="07351-684">Error code</span></span> | <span data-ttu-id="07351-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="07351-686">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="07351-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-687">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-687">Requirements</span></span>

|<span data-ttu-id="07351-688">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-688">Requirement</span></span>| <span data-ttu-id="07351-689">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-690">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-691">1.1</span><span class="sxs-lookup"><span data-stu-id="07351-691">1.1</span></span>|
|[<span data-ttu-id="07351-692">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-694">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-695">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-696">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-696">Example</span></span>

<span data-ttu-id="07351-697">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="07351-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="07351-698">close()</span><span class="sxs-lookup"><span data-stu-id="07351-698">close()</span></span>

<span data-ttu-id="07351-699">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="07351-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="07351-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="07351-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-702">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="07351-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="07351-703">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="07351-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-704">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-704">Requirements</span></span>

|<span data-ttu-id="07351-705">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-705">Requirement</span></span>| <span data-ttu-id="07351-706">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-707">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-708">1.3</span><span class="sxs-lookup"><span data-stu-id="07351-708">1.3</span></span>|
|[<span data-ttu-id="07351-709">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-710">Restrito</span><span class="sxs-lookup"><span data-stu-id="07351-710">Restricted</span></span>|
|[<span data-ttu-id="07351-711">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-712">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="07351-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="07351-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="07351-714">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="07351-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-715">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07351-716">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="07351-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07351-717">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="07351-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="07351-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="07351-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-721">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-721">Parameters:</span></span>

| <span data-ttu-id="07351-722">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-722">Name</span></span> | <span data-ttu-id="07351-723">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-723">Type</span></span> | <span data-ttu-id="07351-724">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-724">Attributes</span></span> | <span data-ttu-id="07351-725">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="07351-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="07351-726">String &#124; Object</span></span>| |<span data-ttu-id="07351-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="07351-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07351-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="07351-729">**OR**</span></span><br/><span data-ttu-id="07351-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="07351-732">String</span><span class="sxs-lookup"><span data-stu-id="07351-732">String</span></span> | <span data-ttu-id="07351-733">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-733">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="07351-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="07351-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="07351-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-737">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-738">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="07351-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="07351-739">String</span><span class="sxs-lookup"><span data-stu-id="07351-739">String</span></span> | | <span data-ttu-id="07351-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="07351-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="07351-742">String</span><span class="sxs-lookup"><span data-stu-id="07351-742">String</span></span> | | <span data-ttu-id="07351-743">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="07351-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="07351-744">String</span><span class="sxs-lookup"><span data-stu-id="07351-744">String</span></span> | | <span data-ttu-id="07351-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07351-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="07351-747">Booliano</span><span class="sxs-lookup"><span data-stu-id="07351-747">Boolean</span></span> | | <span data-ttu-id="07351-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="07351-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="07351-750">String</span><span class="sxs-lookup"><span data-stu-id="07351-750">String</span></span> | | <span data-ttu-id="07351-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="07351-754">function</span><span class="sxs-lookup"><span data-stu-id="07351-754">function</span></span> | <span data-ttu-id="07351-755">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-755">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-756">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-757">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-757">Requirements</span></span>

|<span data-ttu-id="07351-758">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-758">Requirement</span></span>| <span data-ttu-id="07351-759">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-760">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-761">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-761">1.0</span></span>|
|[<span data-ttu-id="07351-762">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-763">ReadItem</span></span>|
|[<span data-ttu-id="07351-764">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-765">Read</span><span class="sxs-lookup"><span data-stu-id="07351-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07351-766">Exemplos</span><span class="sxs-lookup"><span data-stu-id="07351-766">Examples</span></span>

<span data-ttu-id="07351-767">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="07351-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="07351-768">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="07351-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="07351-769">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="07351-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07351-770">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="07351-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07351-771">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="07351-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07351-772">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="07351-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="07351-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="07351-774">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="07351-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-775">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07351-776">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="07351-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07351-777">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="07351-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="07351-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="07351-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-781">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-781">Parameters:</span></span>

| <span data-ttu-id="07351-782">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-782">Name</span></span> | <span data-ttu-id="07351-783">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-783">Type</span></span> | <span data-ttu-id="07351-784">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-784">Attributes</span></span> | <span data-ttu-id="07351-785">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="07351-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="07351-786">String &#124; Object</span></span>| | <span data-ttu-id="07351-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="07351-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07351-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="07351-789">**OR**</span></span><br/><span data-ttu-id="07351-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="07351-792">String</span><span class="sxs-lookup"><span data-stu-id="07351-792">String</span></span> | <span data-ttu-id="07351-793">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-793">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="07351-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="07351-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="07351-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-797">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-798">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="07351-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="07351-799">String</span><span class="sxs-lookup"><span data-stu-id="07351-799">String</span></span> | | <span data-ttu-id="07351-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="07351-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="07351-802">String</span><span class="sxs-lookup"><span data-stu-id="07351-802">String</span></span> | | <span data-ttu-id="07351-803">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="07351-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="07351-804">String</span><span class="sxs-lookup"><span data-stu-id="07351-804">String</span></span> | | <span data-ttu-id="07351-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="07351-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="07351-807">Booliano</span><span class="sxs-lookup"><span data-stu-id="07351-807">Boolean</span></span> | | <span data-ttu-id="07351-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="07351-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="07351-810">String</span><span class="sxs-lookup"><span data-stu-id="07351-810">String</span></span> | | <span data-ttu-id="07351-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="07351-814">function</span><span class="sxs-lookup"><span data-stu-id="07351-814">function</span></span> | <span data-ttu-id="07351-815">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-815">&lt;optional&gt;</span></span> | <span data-ttu-id="07351-816">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-817">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-817">Requirements</span></span>

|<span data-ttu-id="07351-818">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-818">Requirement</span></span>| <span data-ttu-id="07351-819">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-820">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-821">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-821">1.0</span></span>|
|[<span data-ttu-id="07351-822">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-823">ReadItem</span></span>|
|[<span data-ttu-id="07351-824">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-825">Read</span><span class="sxs-lookup"><span data-stu-id="07351-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07351-826">Exemplos</span><span class="sxs-lookup"><span data-stu-id="07351-826">Examples</span></span>

<span data-ttu-id="07351-827">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="07351-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="07351-828">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="07351-828">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="07351-829">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="07351-829">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07351-830">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="07351-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07351-831">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="07351-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07351-832">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="07351-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="07351-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="07351-834">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="07351-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-835">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-836">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-836">Requirements</span></span>

|<span data-ttu-id="07351-837">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-837">Requirement</span></span>| <span data-ttu-id="07351-838">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-839">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-840">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-840">1.0</span></span>|
|[<span data-ttu-id="07351-841">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-842">ReadItem</span></span>|
|[<span data-ttu-id="07351-843">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-844">Read</span><span class="sxs-lookup"><span data-stu-id="07351-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-845">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-845">Returns:</span></span>

<span data-ttu-id="07351-846">Tipo: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="07351-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="07351-847">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-847">Example</span></span>

<span data-ttu-id="07351-848">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-848">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="07351-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="07351-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="07351-850">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="07351-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-851">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-852">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-852">Parameters:</span></span>

|<span data-ttu-id="07351-853">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-853">Name</span></span>| <span data-ttu-id="07351-854">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-854">Type</span></span>| <span data-ttu-id="07351-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="07351-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="07351-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="07351-857">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="07351-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-858">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-858">Requirements</span></span>

|<span data-ttu-id="07351-859">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-859">Requirement</span></span>| <span data-ttu-id="07351-860">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-861">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-862">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-862">1.0</span></span>|
|[<span data-ttu-id="07351-863">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-864">Restrito</span><span class="sxs-lookup"><span data-stu-id="07351-864">Restricted</span></span>|
|[<span data-ttu-id="07351-865">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-866">Read</span><span class="sxs-lookup"><span data-stu-id="07351-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-867">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-867">Returns:</span></span>

<span data-ttu-id="07351-868">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="07351-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="07351-869">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="07351-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="07351-870">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="07351-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="07351-871">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="07351-872">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="07351-872">Value of `entityType`</span></span> | <span data-ttu-id="07351-873">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="07351-873">Type of objects in returned array</span></span> | <span data-ttu-id="07351-874">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="07351-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="07351-875">String</span><span class="sxs-lookup"><span data-stu-id="07351-875">String</span></span> | <span data-ttu-id="07351-876">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="07351-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="07351-877">Contato</span><span class="sxs-lookup"><span data-stu-id="07351-877">Contact</span></span> | <span data-ttu-id="07351-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07351-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="07351-879">String</span><span class="sxs-lookup"><span data-stu-id="07351-879">String</span></span> | <span data-ttu-id="07351-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07351-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="07351-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="07351-881">MeetingSuggestion</span></span> | <span data-ttu-id="07351-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07351-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="07351-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="07351-883">PhoneNumber</span></span> | <span data-ttu-id="07351-884">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="07351-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="07351-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="07351-885">TaskSuggestion</span></span> | <span data-ttu-id="07351-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07351-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="07351-887">String</span><span class="sxs-lookup"><span data-stu-id="07351-887">String</span></span> | <span data-ttu-id="07351-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="07351-888">**Restricted**</span></span> |

<span data-ttu-id="07351-889">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="07351-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="07351-890">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-890">Example</span></span>

<span data-ttu-id="07351-891">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="07351-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="07351-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="07351-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="07351-893">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="07351-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-894">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07351-895">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="07351-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-896">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-896">Parameters:</span></span>

|<span data-ttu-id="07351-897">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-897">Name</span></span>| <span data-ttu-id="07351-898">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-898">Type</span></span>| <span data-ttu-id="07351-899">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="07351-900">String</span><span class="sxs-lookup"><span data-stu-id="07351-900">String</span></span>|<span data-ttu-id="07351-901">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="07351-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-902">Requirements</span></span>

|<span data-ttu-id="07351-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-903">Requirement</span></span>| <span data-ttu-id="07351-904">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-905">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-906">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-906">1.0</span></span>|
|[<span data-ttu-id="07351-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-908">ReadItem</span></span>|
|[<span data-ttu-id="07351-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-910">Read</span><span class="sxs-lookup"><span data-stu-id="07351-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-911">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-911">Returns:</span></span>

<span data-ttu-id="07351-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="07351-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="07351-914">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="07351-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="07351-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="07351-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="07351-916">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="07351-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-917">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07351-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="07351-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="07351-921">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="07351-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="07351-922">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="07351-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="07351-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="07351-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07351-926">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-926">Requirements</span></span>

|<span data-ttu-id="07351-927">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-927">Requirement</span></span>| <span data-ttu-id="07351-928">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-929">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-930">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-930">1.0</span></span>|
|[<span data-ttu-id="07351-931">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-932">ReadItem</span></span>|
|[<span data-ttu-id="07351-933">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-934">Read</span><span class="sxs-lookup"><span data-stu-id="07351-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-935">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-935">Returns:</span></span>

<span data-ttu-id="07351-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="07351-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="07351-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="07351-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07351-939">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07351-940">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-940">Example</span></span>

<span data-ttu-id="07351-941">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="07351-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="07351-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="07351-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="07351-943">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="07351-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-944">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="07351-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07351-945">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="07351-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="07351-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="07351-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-948">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-948">Parameters:</span></span>

|<span data-ttu-id="07351-949">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-949">Name</span></span>| <span data-ttu-id="07351-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-950">Type</span></span>| <span data-ttu-id="07351-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="07351-952">String</span><span class="sxs-lookup"><span data-stu-id="07351-952">String</span></span>|<span data-ttu-id="07351-953">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="07351-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-954">Requirements</span></span>

|<span data-ttu-id="07351-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-955">Requirement</span></span>| <span data-ttu-id="07351-956">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-957">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-958">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-958">1.0</span></span>|
|[<span data-ttu-id="07351-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-960">ReadItem</span></span>|
|[<span data-ttu-id="07351-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-962">Read</span><span class="sxs-lookup"><span data-stu-id="07351-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-963">Returns:</span></span>

<span data-ttu-id="07351-964">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="07351-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="07351-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="07351-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07351-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="07351-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07351-967">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-967">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="07351-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="07351-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="07351-969">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="07351-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="07351-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-972">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-972">Parameters:</span></span>

|<span data-ttu-id="07351-973">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-973">Name</span></span>| <span data-ttu-id="07351-974">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-974">Type</span></span>| <span data-ttu-id="07351-975">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-975">Attributes</span></span>| <span data-ttu-id="07351-976">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="07351-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07351-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="07351-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="07351-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="07351-981">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-981">Object</span></span>| <span data-ttu-id="07351-982">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-982">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-983">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07351-984">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-984">Object</span></span>| <span data-ttu-id="07351-985">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-985">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-986">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07351-987">function</span><span class="sxs-lookup"><span data-stu-id="07351-987">function</span></span>||<span data-ttu-id="07351-988">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07351-989">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="07351-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="07351-990">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="07351-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-991">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-991">Requirements</span></span>

|<span data-ttu-id="07351-992">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-992">Requirement</span></span>| <span data-ttu-id="07351-993">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-994">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-995">1.2</span><span class="sxs-lookup"><span data-stu-id="07351-995">1.2</span></span>|
|[<span data-ttu-id="07351-996">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-998">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-999">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="07351-1000">Retorna:</span><span class="sxs-lookup"><span data-stu-id="07351-1000">Returns:</span></span>

<span data-ttu-id="07351-1001">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="07351-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="07351-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="07351-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07351-1003">String</span><span class="sxs-lookup"><span data-stu-id="07351-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07351-1004">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-1004">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="07351-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="07351-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="07351-1006">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="07351-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="07351-p163">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="07351-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-1010">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-1010">Parameters:</span></span>

|<span data-ttu-id="07351-1011">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-1011">Name</span></span>| <span data-ttu-id="07351-1012">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-1012">Type</span></span>| <span data-ttu-id="07351-1013">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-1013">Attributes</span></span>| <span data-ttu-id="07351-1014">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="07351-1015">function</span><span class="sxs-lookup"><span data-stu-id="07351-1015">function</span></span>||<span data-ttu-id="07351-1016">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07351-1017">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="07351-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="07351-1018">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="07351-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="07351-1019">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1019">Object</span></span>| <span data-ttu-id="07351-1020">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1021">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="07351-1022">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-1023">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-1023">Requirements</span></span>

|<span data-ttu-id="07351-1024">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-1024">Requirement</span></span>| <span data-ttu-id="07351-1025">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-1026">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="07351-1027">1.0</span></span>|
|[<span data-ttu-id="07351-1028">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07351-1029">ReadItem</span></span>|
|[<span data-ttu-id="07351-1030">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-1031">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="07351-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-1032">Example</span></span>

<span data-ttu-id="07351-p166">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="07351-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="07351-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07351-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="07351-1037">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="07351-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="07351-p167">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="07351-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-1042">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-1042">Parameters:</span></span>

|<span data-ttu-id="07351-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-1043">Name</span></span>| <span data-ttu-id="07351-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-1044">Type</span></span>| <span data-ttu-id="07351-1045">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-1045">Attributes</span></span>| <span data-ttu-id="07351-1046">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="07351-1047">String</span><span class="sxs-lookup"><span data-stu-id="07351-1047">String</span></span>||<span data-ttu-id="07351-p168">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="07351-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="07351-1050">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1050">Object</span></span>| <span data-ttu-id="07351-1051">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1052">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07351-1053">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1053">Object</span></span>| <span data-ttu-id="07351-1054">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1055">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07351-1056">function</span><span class="sxs-lookup"><span data-stu-id="07351-1056">function</span></span>| <span data-ttu-id="07351-1057">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1058">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07351-1059">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="07351-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07351-1060">Erros</span><span class="sxs-lookup"><span data-stu-id="07351-1060">Errors</span></span>

| <span data-ttu-id="07351-1061">Código de erro</span><span class="sxs-lookup"><span data-stu-id="07351-1061">Error code</span></span> | <span data-ttu-id="07351-1062">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="07351-1063">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="07351-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-1064">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-1064">Requirements</span></span>

|<span data-ttu-id="07351-1065">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-1065">Requirement</span></span>| <span data-ttu-id="07351-1066">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-1067">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="07351-1068">1.1</span></span>|
|[<span data-ttu-id="07351-1069">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-1071">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-1072">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-1073">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-1073">Example</span></span>

<span data-ttu-id="07351-1074">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="07351-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="07351-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="07351-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="07351-1076">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="07351-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="07351-p169">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="07351-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-1080">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="07351-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="07351-1081">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="07351-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="07351-p171">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="07351-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="07351-1085">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="07351-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="07351-1086">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="07351-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="07351-1087">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="07351-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="07351-1088">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="07351-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-1089">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-1089">Parameters:</span></span>

|<span data-ttu-id="07351-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-1090">Name</span></span>| <span data-ttu-id="07351-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-1091">Type</span></span>| <span data-ttu-id="07351-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-1092">Attributes</span></span>| <span data-ttu-id="07351-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="07351-1094">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1094">Object</span></span>| <span data-ttu-id="07351-1095">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1096">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07351-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1097">Object</span></span>| <span data-ttu-id="07351-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1099">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="07351-1100">function</span><span class="sxs-lookup"><span data-stu-id="07351-1100">function</span></span>||<span data-ttu-id="07351-1101">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07351-1102">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="07351-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07351-1103">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-1103">Requirements</span></span>

|<span data-ttu-id="07351-1104">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-1104">Requirement</span></span>| <span data-ttu-id="07351-1105">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-1106">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="07351-1107">1.3</span></span>|
|[<span data-ttu-id="07351-1108">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-1110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-1111">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07351-1112">Exemplos</span><span class="sxs-lookup"><span data-stu-id="07351-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="07351-p173">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="07351-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="07351-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="07351-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="07351-1116">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="07351-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="07351-p174">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="07351-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07351-1120">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="07351-1120">Parameters:</span></span>

|<span data-ttu-id="07351-1121">Nome</span><span class="sxs-lookup"><span data-stu-id="07351-1121">Name</span></span>| <span data-ttu-id="07351-1122">Tipo</span><span class="sxs-lookup"><span data-stu-id="07351-1122">Type</span></span>| <span data-ttu-id="07351-1123">Atributos</span><span class="sxs-lookup"><span data-stu-id="07351-1123">Attributes</span></span>| <span data-ttu-id="07351-1124">Descrição</span><span class="sxs-lookup"><span data-stu-id="07351-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="07351-1125">String</span><span class="sxs-lookup"><span data-stu-id="07351-1125">String</span></span>||<span data-ttu-id="07351-p175">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="07351-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="07351-1129">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1129">Object</span></span>| <span data-ttu-id="07351-1130">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1131">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="07351-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="07351-1132">Objeto</span><span class="sxs-lookup"><span data-stu-id="07351-1132">Object</span></span>| <span data-ttu-id="07351-1133">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-1134">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="07351-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="07351-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07351-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="07351-1136">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="07351-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="07351-p176">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="07351-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="07351-p177">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="07351-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="07351-1141">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="07351-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="07351-1142">function</span><span class="sxs-lookup"><span data-stu-id="07351-1142">function</span></span>||<span data-ttu-id="07351-1143">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="07351-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="07351-1144">Requisitos</span><span class="sxs-lookup"><span data-stu-id="07351-1144">Requirements</span></span>

|<span data-ttu-id="07351-1145">Requisito</span><span class="sxs-lookup"><span data-stu-id="07351-1145">Requirement</span></span>| <span data-ttu-id="07351-1146">Valor</span><span class="sxs-lookup"><span data-stu-id="07351-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="07351-1147">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="07351-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07351-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="07351-1148">1.2</span></span>|
|[<span data-ttu-id="07351-1149">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="07351-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07351-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07351-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="07351-1151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="07351-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="07351-1152">Escrever</span><span class="sxs-lookup"><span data-stu-id="07351-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07351-1153">Exemplo</span><span class="sxs-lookup"><span data-stu-id="07351-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```