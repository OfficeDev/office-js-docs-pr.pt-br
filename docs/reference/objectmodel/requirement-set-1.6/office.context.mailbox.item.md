
# <a name="item"></a><span data-ttu-id="ee128-101">item</span><span class="sxs-lookup"><span data-stu-id="ee128-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ee128-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ee128-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ee128-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ee128-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-105">Requirements</span></span>

|<span data-ttu-id="ee128-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-106">Requirement</span></span>| <span data-ttu-id="ee128-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-109">1.0</span></span>|
|[<span data-ttu-id="ee128-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="ee128-111">Restricted</span></span>|
|[<span data-ttu-id="ee128-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ee128-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ee128-114">Members and methods</span></span>

| <span data-ttu-id="ee128-115">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-115">Member</span></span> | <span data-ttu-id="ee128-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ee128-117">attachments</span><span class="sxs-lookup"><span data-stu-id="ee128-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="ee128-118">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-118">Member</span></span> |
| [<span data-ttu-id="ee128-119">bcc</span><span class="sxs-lookup"><span data-stu-id="ee128-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ee128-120">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-120">Member</span></span> |
| [<span data-ttu-id="ee128-121">body</span><span class="sxs-lookup"><span data-stu-id="ee128-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="ee128-122">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-122">Member</span></span> |
| [<span data-ttu-id="ee128-123">cc</span><span class="sxs-lookup"><span data-stu-id="ee128-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ee128-124">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-124">Member</span></span> |
| [<span data-ttu-id="ee128-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="ee128-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ee128-126">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-126">Member</span></span> |
| [<span data-ttu-id="ee128-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ee128-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ee128-128">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-128">Member</span></span> |
| [<span data-ttu-id="ee128-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ee128-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ee128-130">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-130">Member</span></span> |
| [<span data-ttu-id="ee128-131">end</span><span class="sxs-lookup"><span data-stu-id="ee128-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="ee128-132">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-132">Member</span></span> |
| [<span data-ttu-id="ee128-133">from</span><span class="sxs-lookup"><span data-stu-id="ee128-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ee128-134">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-134">Member</span></span> |
| [<span data-ttu-id="ee128-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ee128-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ee128-136">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-136">Member</span></span> |
| [<span data-ttu-id="ee128-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="ee128-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ee128-138">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-138">Member</span></span> |
| [<span data-ttu-id="ee128-139">itemId</span><span class="sxs-lookup"><span data-stu-id="ee128-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ee128-140">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-140">Member</span></span> |
| [<span data-ttu-id="ee128-141">itemType</span><span class="sxs-lookup"><span data-stu-id="ee128-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="ee128-142">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-142">Member</span></span> |
| [<span data-ttu-id="ee128-143">location</span><span class="sxs-lookup"><span data-stu-id="ee128-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="ee128-144">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-144">Member</span></span> |
| [<span data-ttu-id="ee128-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ee128-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ee128-146">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-146">Member</span></span> |
| [<span data-ttu-id="ee128-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ee128-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="ee128-148">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-148">Member</span></span> |
| [<span data-ttu-id="ee128-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ee128-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ee128-150">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-150">Member</span></span> |
| [<span data-ttu-id="ee128-151">organizer</span><span class="sxs-lookup"><span data-stu-id="ee128-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ee128-152">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-152">Member</span></span> |
| [<span data-ttu-id="ee128-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ee128-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ee128-154">Member</span><span class="sxs-lookup"><span data-stu-id="ee128-154">Member</span></span> |
| [<span data-ttu-id="ee128-155">sender</span><span class="sxs-lookup"><span data-stu-id="ee128-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ee128-156">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-156">Member</span></span> |
| [<span data-ttu-id="ee128-157">start</span><span class="sxs-lookup"><span data-stu-id="ee128-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="ee128-158">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-158">Member</span></span> |
| [<span data-ttu-id="ee128-159">subject</span><span class="sxs-lookup"><span data-stu-id="ee128-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="ee128-160">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-160">Member</span></span> |
| [<span data-ttu-id="ee128-161">to</span><span class="sxs-lookup"><span data-stu-id="ee128-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ee128-162">Membro</span><span class="sxs-lookup"><span data-stu-id="ee128-162">Member</span></span> |
| [<span data-ttu-id="ee128-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ee128-164">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-164">Method</span></span> |
| [<span data-ttu-id="ee128-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ee128-166">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-166">Method</span></span> |
| [<span data-ttu-id="ee128-167">close</span><span class="sxs-lookup"><span data-stu-id="ee128-167">close</span></span>](#close) | <span data-ttu-id="ee128-168">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-168">Method</span></span> |
| [<span data-ttu-id="ee128-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ee128-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="ee128-170">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-170">Method</span></span> |
| [<span data-ttu-id="ee128-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ee128-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="ee128-172">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-172">Method</span></span> |
| [<span data-ttu-id="ee128-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="ee128-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="ee128-174">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-174">Method</span></span> |
| [<span data-ttu-id="ee128-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ee128-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="ee128-176">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-176">Method</span></span> |
| [<span data-ttu-id="ee128-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ee128-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="ee128-178">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-178">Method</span></span> |
| [<span data-ttu-id="ee128-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ee128-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ee128-180">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-180">Method</span></span> |
| [<span data-ttu-id="ee128-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ee128-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ee128-182">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-182">Method</span></span> |
| [<span data-ttu-id="ee128-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ee128-184">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-184">Method</span></span> |
| [<span data-ttu-id="ee128-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="ee128-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="ee128-186">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-186">Method</span></span> |
| [<span data-ttu-id="ee128-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ee128-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ee128-188">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-188">Method</span></span> |
| [<span data-ttu-id="ee128-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ee128-190">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-190">Method</span></span> |
| [<span data-ttu-id="ee128-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ee128-192">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-192">Method</span></span> |
| [<span data-ttu-id="ee128-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ee128-194">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-194">Method</span></span> |
| [<span data-ttu-id="ee128-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ee128-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ee128-196">Método</span><span class="sxs-lookup"><span data-stu-id="ee128-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ee128-197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-197">Example</span></span>

<span data-ttu-id="ee128-198">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee128-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ee128-199">Membros</span><span class="sxs-lookup"><span data-stu-id="ee128-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="ee128-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ee128-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="ee128-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-203">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="ee128-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ee128-204">Para saber mais, consulte [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ee128-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-205">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-205">Type:</span></span>

*   <span data-ttu-id="ee128-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ee128-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-207">Requirements</span></span>

|<span data-ttu-id="ee128-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-208">Requirement</span></span>| <span data-ttu-id="ee128-209">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-211">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-211">1.0</span></span>|
|[<span data-ttu-id="ee128-212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-213">ReadItem</span></span>|
|[<span data-ttu-id="ee128-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-215">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-216">Example</span></span>

<span data-ttu-id="ee128-217">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ee128-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ee128-219">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ee128-220">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ee128-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-221">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-221">Type:</span></span>

*   [<span data-ttu-id="ee128-222">Destinatários</span><span class="sxs-lookup"><span data-stu-id="ee128-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ee128-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-223">Requirements</span></span>

|<span data-ttu-id="ee128-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-224">Requirement</span></span>| <span data-ttu-id="ee128-225">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-226">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-227">1.1</span><span class="sxs-lookup"><span data-stu-id="ee128-227">1.1</span></span>|
|[<span data-ttu-id="ee128-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-229">ReadItem</span></span>|
|[<span data-ttu-id="ee128-230">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-231">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-232">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-232">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="ee128-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="ee128-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="ee128-234">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="ee128-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-235">Type:</span></span>

*   [<span data-ttu-id="ee128-236">Corpo</span><span class="sxs-lookup"><span data-stu-id="ee128-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="ee128-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-237">Requirements</span></span>

|<span data-ttu-id="ee128-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-238">Requirement</span></span>| <span data-ttu-id="ee128-239">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-241">1.1</span><span class="sxs-lookup"><span data-stu-id="ee128-241">1.1</span></span>|
|[<span data-ttu-id="ee128-242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-243">ReadItem</span></span>|
|[<span data-ttu-id="ee128-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-245">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ee128-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ee128-247">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ee128-248">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-249">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-249">Read mode</span></span>

<span data-ttu-id="ee128-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="ee128-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-252">Compose mode</span></span>

<span data-ttu-id="ee128-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-254">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-254">Type:</span></span>

*   <span data-ttu-id="ee128-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-256">Requirements</span></span>

|<span data-ttu-id="ee128-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-257">Requirement</span></span>| <span data-ttu-id="ee128-258">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-260">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-260">1.0</span></span>|
|[<span data-ttu-id="ee128-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-262">ReadItem</span></span>|
|[<span data-ttu-id="ee128-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-264">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-265">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="ee128-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="ee128-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="ee128-267">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="ee128-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ee128-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="ee128-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ee128-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="ee128-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-272">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-272">Type:</span></span>

*   <span data-ttu-id="ee128-273">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee128-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-274">Requirements</span></span>

|<span data-ttu-id="ee128-275">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-275">Requirement</span></span>| <span data-ttu-id="ee128-276">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-277">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-278">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-278">1.0</span></span>|
|[<span data-ttu-id="ee128-279">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-280">ReadItem</span></span>|
|[<span data-ttu-id="ee128-281">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-282">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="ee128-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="ee128-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="ee128-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-286">Type:</span></span>

*   <span data-ttu-id="ee128-287">Data</span><span class="sxs-lookup"><span data-stu-id="ee128-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-288">Requirements</span></span>

|<span data-ttu-id="ee128-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-289">Requirement</span></span>| <span data-ttu-id="ee128-290">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-292">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-292">1.0</span></span>|
|[<span data-ttu-id="ee128-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-294">ReadItem</span></span>|
|[<span data-ttu-id="ee128-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-296">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-297">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ee128-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="ee128-298">dateTimeModified :Date</span></span>

<span data-ttu-id="ee128-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-301">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-302">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-302">Type:</span></span>

*   <span data-ttu-id="ee128-303">Data</span><span class="sxs-lookup"><span data-stu-id="ee128-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-304">Requirements</span></span>

|<span data-ttu-id="ee128-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-305">Requirement</span></span>| <span data-ttu-id="ee128-306">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-307">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-308">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-308">1.0</span></span>|
|[<span data-ttu-id="ee128-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-310">ReadItem</span></span>|
|[<span data-ttu-id="ee128-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-312">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-313">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="ee128-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ee128-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="ee128-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="ee128-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ee128-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="ee128-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-318">Read mode</span></span>

<span data-ttu-id="ee128-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ee128-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-320">Compose mode</span></span>

<span data-ttu-id="ee128-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ee128-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ee128-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="ee128-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-323">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-323">Type:</span></span>

*   <span data-ttu-id="ee128-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ee128-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-325">Requirements</span></span>

|<span data-ttu-id="ee128-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-326">Requirement</span></span>| <span data-ttu-id="ee128-327">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-328">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-329">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-329">1.0</span></span>|
|[<span data-ttu-id="ee128-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-331">ReadItem</span></span>|
|[<span data-ttu-id="ee128-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-333">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-334">Example</span></span>

<span data-ttu-id="ee128-335">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ee128-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ee128-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ee128-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ee128-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ee128-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="ee128-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-341">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ee128-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-342">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-342">Type:</span></span>

*   [<span data-ttu-id="ee128-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ee128-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ee128-344">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-344">Requirements</span></span>

|<span data-ttu-id="ee128-345">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-345">Requirement</span></span>| <span data-ttu-id="ee128-346">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-347">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-347">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-348">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-348">1.0</span></span>|
|[<span data-ttu-id="ee128-349">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-350">ReadItem</span></span>|
|[<span data-ttu-id="ee128-351">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-352">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="ee128-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="ee128-353">internetMessageId :String</span></span>

<span data-ttu-id="ee128-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-356">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-356">Type:</span></span>

*   <span data-ttu-id="ee128-357">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee128-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-358">Requirements</span></span>

|<span data-ttu-id="ee128-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-359">Requirement</span></span>| <span data-ttu-id="ee128-360">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-361">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-362">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-362">1.0</span></span>|
|[<span data-ttu-id="ee128-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-364">ReadItem</span></span>|
|[<span data-ttu-id="ee128-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-366">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ee128-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="ee128-368">itemClass :String</span></span>

<span data-ttu-id="ee128-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ee128-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ee128-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-373">Type</span></span> | <span data-ttu-id="ee128-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-374">Description</span></span> | <span data-ttu-id="ee128-375">classe de item</span><span class="sxs-lookup"><span data-stu-id="ee128-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ee128-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="ee128-376">Appointment items</span></span> | <span data-ttu-id="ee128-377">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="ee128-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="ee128-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="ee128-378">Message items</span></span> | <span data-ttu-id="ee128-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="ee128-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ee128-380">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee128-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-381">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-381">Type:</span></span>

*   <span data-ttu-id="ee128-382">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee128-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-383">Requirements</span></span>

|<span data-ttu-id="ee128-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-384">Requirement</span></span>| <span data-ttu-id="ee128-385">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-386">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-387">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-387">1.0</span></span>|
|[<span data-ttu-id="ee128-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-389">ReadItem</span></span>|
|[<span data-ttu-id="ee128-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-391">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ee128-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="ee128-393">(nullable) itemId :String</span></span>

<span data-ttu-id="ee128-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="ee128-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ee128-397">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee128-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ee128-398">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="ee128-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ee128-399">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ee128-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ee128-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-402">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-402">Type:</span></span>

*   <span data-ttu-id="ee128-403">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee128-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-404">Requirements</span></span>

|<span data-ttu-id="ee128-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-405">Requirement</span></span>| <span data-ttu-id="ee128-406">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-407">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-408">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-408">1.0</span></span>|
|[<span data-ttu-id="ee128-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-410">ReadItem</span></span>|
|[<span data-ttu-id="ee128-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-412">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-413">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-413">Example</span></span>

<span data-ttu-id="ee128-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="ee128-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="ee128-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ee128-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ee128-417">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="ee128-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ee128-418">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-419">Type:</span></span>

*   [<span data-ttu-id="ee128-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ee128-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ee128-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-421">Requirements</span></span>

|<span data-ttu-id="ee128-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-422">Requirement</span></span>| <span data-ttu-id="ee128-423">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-425">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-425">1.0</span></span>|
|[<span data-ttu-id="ee128-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-427">ReadItem</span></span>|
|[<span data-ttu-id="ee128-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-429">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-430">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="ee128-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="ee128-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="ee128-432">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-433">Read mode</span></span>

<span data-ttu-id="ee128-434">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-435">Compose mode</span></span>

<span data-ttu-id="ee128-436">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-437">Type:</span></span>

*   <span data-ttu-id="ee128-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="ee128-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-439">Requirements</span></span>

|<span data-ttu-id="ee128-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-440">Requirement</span></span>| <span data-ttu-id="ee128-441">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-443">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-443">1.0</span></span>|
|[<span data-ttu-id="ee128-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-445">ReadItem</span></span>|
|[<span data-ttu-id="ee128-446">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-447">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-448">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ee128-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="ee128-449">normalizedSubject :String</span></span>

<span data-ttu-id="ee128-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ee128-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="ee128-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-454">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-454">Type:</span></span>

*   <span data-ttu-id="ee128-455">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee128-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-456">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-456">Requirements</span></span>

|<span data-ttu-id="ee128-457">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-457">Requirement</span></span>| <span data-ttu-id="ee128-458">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-459">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-459">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-460">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-460">1.0</span></span>|
|[<span data-ttu-id="ee128-461">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-462">ReadItem</span></span>|
|[<span data-ttu-id="ee128-463">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-464">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-465">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-465">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="ee128-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="ee128-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="ee128-467">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="ee128-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-468">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-468">Type:</span></span>

*   [<span data-ttu-id="ee128-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ee128-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="ee128-470">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-470">Requirements</span></span>

|<span data-ttu-id="ee128-471">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-471">Requirement</span></span>| <span data-ttu-id="ee128-472">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-473">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-474">1.3</span><span class="sxs-lookup"><span data-stu-id="ee128-474">1.3</span></span>|
|[<span data-ttu-id="ee128-475">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-476">ReadItem</span></span>|
|[<span data-ttu-id="ee128-477">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-478">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ee128-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ee128-480">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="ee128-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ee128-481">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-482">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-482">Read mode</span></span>

<span data-ttu-id="ee128-483">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="ee128-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-484">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-484">Compose mode</span></span>

<span data-ttu-id="ee128-485">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="ee128-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-486">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-486">Type:</span></span>

*   <span data-ttu-id="ee128-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-488">Requirements</span></span>

|<span data-ttu-id="ee128-489">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-489">Requirement</span></span>| <span data-ttu-id="ee128-490">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-491">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-492">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-492">1.0</span></span>|
|[<span data-ttu-id="ee128-493">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-494">ReadItem</span></span>|
|[<span data-ttu-id="ee128-495">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-496">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-497">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-497">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ee128-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ee128-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ee128-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-501">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-501">Type:</span></span>

*   [<span data-ttu-id="ee128-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ee128-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ee128-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-503">Requirements</span></span>

|<span data-ttu-id="ee128-504">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-504">Requirement</span></span>| <span data-ttu-id="ee128-505">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-506">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-507">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-507">1.0</span></span>|
|[<span data-ttu-id="ee128-508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-509">ReadItem</span></span>|
|[<span data-ttu-id="ee128-510">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-511">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-512">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-512">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ee128-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ee128-514">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="ee128-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ee128-515">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-516">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-516">Read mode</span></span>

<span data-ttu-id="ee128-517">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="ee128-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-518">Compose mode</span></span>

<span data-ttu-id="ee128-519">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="ee128-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-520">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-520">Type:</span></span>

*   <span data-ttu-id="ee128-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-522">Requirements</span></span>

|<span data-ttu-id="ee128-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-523">Requirement</span></span>| <span data-ttu-id="ee128-524">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-526">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-526">1.0</span></span>|
|[<span data-ttu-id="ee128-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-528">ReadItem</span></span>|
|[<span data-ttu-id="ee128-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-530">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-531">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-531">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ee128-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ee128-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ee128-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ee128-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ee128-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="ee128-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-537">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ee128-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-538">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-538">Type:</span></span>

*   [<span data-ttu-id="ee128-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ee128-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ee128-540">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-540">Requirements</span></span>

|<span data-ttu-id="ee128-541">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-541">Requirement</span></span>| <span data-ttu-id="ee128-542">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-543">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-544">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-544">1.0</span></span>|
|[<span data-ttu-id="ee128-545">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-546">ReadItem</span></span>|
|[<span data-ttu-id="ee128-547">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-548">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-549">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-549">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="ee128-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ee128-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="ee128-551">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="ee128-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ee128-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="ee128-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-554">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-554">Read mode</span></span>

<span data-ttu-id="ee128-555">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ee128-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-556">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-556">Compose mode</span></span>

<span data-ttu-id="ee128-557">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ee128-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ee128-558">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="ee128-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-559">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-559">Type:</span></span>

*   <span data-ttu-id="ee128-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ee128-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-561">Requirements</span></span>

|<span data-ttu-id="ee128-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-562">Requirement</span></span>| <span data-ttu-id="ee128-563">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-565">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-565">1.0</span></span>|
|[<span data-ttu-id="ee128-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-567">ReadItem</span></span>|
|[<span data-ttu-id="ee128-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-569">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-570">Example</span></span>

<span data-ttu-id="ee128-571">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ee128-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="ee128-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ee128-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="ee128-573">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="ee128-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ee128-574">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="ee128-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-575">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-575">Read mode</span></span>

<span data-ttu-id="ee128-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ee128-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="ee128-578">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-578">Compose mode</span></span>

<span data-ttu-id="ee128-579">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="ee128-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ee128-580">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-580">Type:</span></span>

*   <span data-ttu-id="ee128-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ee128-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-582">Requirements</span></span>

|<span data-ttu-id="ee128-583">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-583">Requirement</span></span>| <span data-ttu-id="ee128-584">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-585">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-586">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-586">1.0</span></span>|
|[<span data-ttu-id="ee128-587">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-588">ReadItem</span></span>|
|[<span data-ttu-id="ee128-589">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-590">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ee128-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ee128-592">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ee128-593">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ee128-594">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-594">Read mode</span></span>

<span data-ttu-id="ee128-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="ee128-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ee128-597">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ee128-597">Compose mode</span></span>

<span data-ttu-id="ee128-598">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ee128-599">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee128-599">Type:</span></span>

*   <span data-ttu-id="ee128-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ee128-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-601">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-601">Requirements</span></span>

|<span data-ttu-id="ee128-602">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-602">Requirement</span></span>| <span data-ttu-id="ee128-603">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-604">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-605">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-605">1.0</span></span>|
|[<span data-ttu-id="ee128-606">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-607">ReadItem</span></span>|
|[<span data-ttu-id="ee128-608">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-609">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-610">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-610">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="ee128-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="ee128-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ee128-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ee128-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ee128-613">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="ee128-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ee128-614">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="ee128-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ee128-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="ee128-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-616">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-616">Parameters:</span></span>

|<span data-ttu-id="ee128-617">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-617">Name</span></span>| <span data-ttu-id="ee128-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-618">Type</span></span>| <span data-ttu-id="ee128-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-619">Attributes</span></span>| <span data-ttu-id="ee128-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ee128-621">String</span><span class="sxs-lookup"><span data-stu-id="ee128-621">String</span></span>||<span data-ttu-id="ee128-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ee128-624">String</span><span class="sxs-lookup"><span data-stu-id="ee128-624">String</span></span>||<span data-ttu-id="ee128-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ee128-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-627">Object</span></span>| <span data-ttu-id="ee128-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-628">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-629">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="ee128-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-630">Object</span></span> | <span data-ttu-id="ee128-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-631">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="ee128-633">Booliano</span><span class="sxs-lookup"><span data-stu-id="ee128-633">Boolean</span></span> | <span data-ttu-id="ee128-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-634">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-635">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ee128-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="ee128-636">function</span><span class="sxs-lookup"><span data-stu-id="ee128-636">function</span></span>| <span data-ttu-id="ee128-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-637">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-638">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ee128-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ee128-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ee128-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="ee128-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ee128-641">Erros</span><span class="sxs-lookup"><span data-stu-id="ee128-641">Errors</span></span>

| <span data-ttu-id="ee128-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ee128-642">Error code</span></span> | <span data-ttu-id="ee128-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ee128-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="ee128-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ee128-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="ee128-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ee128-646">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="ee128-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-647">Requirements</span></span>

|<span data-ttu-id="ee128-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-648">Requirement</span></span>| <span data-ttu-id="ee128-649">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-650">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-651">1.1</span><span class="sxs-lookup"><span data-stu-id="ee128-651">1.1</span></span>|
|[<span data-ttu-id="ee128-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-655">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ee128-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ee128-656">Examples</span></span>

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

<span data-ttu-id="ee128-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ee128-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ee128-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ee128-659">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ee128-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="ee128-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ee128-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="ee128-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ee128-664">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="ee128-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-665">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-665">Parameters:</span></span>

|<span data-ttu-id="ee128-666">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-666">Name</span></span>| <span data-ttu-id="ee128-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-667">Type</span></span>| <span data-ttu-id="ee128-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-668">Attributes</span></span>| <span data-ttu-id="ee128-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ee128-670">String</span><span class="sxs-lookup"><span data-stu-id="ee128-670">String</span></span>||<span data-ttu-id="ee128-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ee128-673">String</span><span class="sxs-lookup"><span data-stu-id="ee128-673">String</span></span>||<span data-ttu-id="ee128-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ee128-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-676">Object</span></span>| <span data-ttu-id="ee128-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-677">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ee128-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-679">Object</span></span>| <span data-ttu-id="ee128-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-680">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ee128-682">function</span><span class="sxs-lookup"><span data-stu-id="ee128-682">function</span></span>| <span data-ttu-id="ee128-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-683">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ee128-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ee128-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ee128-686">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="ee128-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ee128-687">Erros</span><span class="sxs-lookup"><span data-stu-id="ee128-687">Errors</span></span>

| <span data-ttu-id="ee128-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ee128-688">Error code</span></span> | <span data-ttu-id="ee128-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ee128-690">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="ee128-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-691">Requirements</span></span>

|<span data-ttu-id="ee128-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-692">Requirement</span></span>| <span data-ttu-id="ee128-693">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-694">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-695">1.1</span><span class="sxs-lookup"><span data-stu-id="ee128-695">1.1</span></span>|
|[<span data-ttu-id="ee128-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-699">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-700">Example</span></span>

<span data-ttu-id="ee128-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ee128-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="ee128-702">close()</span><span class="sxs-lookup"><span data-stu-id="ee128-702">close()</span></span>

<span data-ttu-id="ee128-703">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="ee128-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ee128-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="ee128-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-706">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="ee128-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ee128-707">No cliente do Outlook para área de trabalho, se a mensagem é uma resposta embutida, o método `close` não tem efeito.</span><span class="sxs-lookup"><span data-stu-id="ee128-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-708">Requirements</span></span>

|<span data-ttu-id="ee128-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-709">Requirement</span></span>| <span data-ttu-id="ee128-710">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-711">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-712">1.3</span><span class="sxs-lookup"><span data-stu-id="ee128-712">1.3</span></span>|
|[<span data-ttu-id="ee128-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="ee128-714">Restricted</span></span>|
|[<span data-ttu-id="ee128-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-716">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="ee128-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ee128-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="ee128-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="ee128-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-719">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-720">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="ee128-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ee128-721">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="ee128-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ee128-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="ee128-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-725">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-725">Parameters:</span></span>

| <span data-ttu-id="ee128-726">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-726">Name</span></span> | <span data-ttu-id="ee128-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-727">Type</span></span> | <span data-ttu-id="ee128-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-728">Attributes</span></span> | <span data-ttu-id="ee128-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ee128-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ee128-730">String &#124; Object</span></span>| |<span data-ttu-id="ee128-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ee128-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ee128-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="ee128-733">**OR**</span></span><br/><span data-ttu-id="ee128-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ee128-736">String</span><span class="sxs-lookup"><span data-stu-id="ee128-736">String</span></span> | <span data-ttu-id="ee128-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-737">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ee128-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ee128-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ee128-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-741">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-742">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="ee128-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ee128-743">String</span><span class="sxs-lookup"><span data-stu-id="ee128-743">String</span></span> | | <span data-ttu-id="ee128-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ee128-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ee128-746">String</span><span class="sxs-lookup"><span data-stu-id="ee128-746">String</span></span> | | <span data-ttu-id="ee128-747">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="ee128-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ee128-748">String</span><span class="sxs-lookup"><span data-stu-id="ee128-748">String</span></span> | | <span data-ttu-id="ee128-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ee128-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ee128-751">Booliano</span><span class="sxs-lookup"><span data-stu-id="ee128-751">Boolean</span></span> | | <span data-ttu-id="ee128-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ee128-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ee128-754">String</span><span class="sxs-lookup"><span data-stu-id="ee128-754">String</span></span> | | <span data-ttu-id="ee128-p145">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ee128-758">function</span><span class="sxs-lookup"><span data-stu-id="ee128-758">function</span></span> | <span data-ttu-id="ee128-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-759">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-760">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-761">Requirements</span></span>

|<span data-ttu-id="ee128-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-762">Requirement</span></span>| <span data-ttu-id="ee128-763">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-764">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-765">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-765">1.0</span></span>|
|[<span data-ttu-id="ee128-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-767">ReadItem</span></span>|
|[<span data-ttu-id="ee128-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-769">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ee128-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ee128-770">Examples</span></span>

<span data-ttu-id="ee128-771">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ee128-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ee128-772">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="ee128-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ee128-773">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="ee128-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ee128-774">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="ee128-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ee128-775">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ee128-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ee128-776">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="ee128-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ee128-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="ee128-778">Exibe um formulário de resposta que inclui somente o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="ee128-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-779">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-780">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="ee128-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ee128-781">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="ee128-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ee128-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="ee128-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-785">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-785">Parameters:</span></span>

| <span data-ttu-id="ee128-786">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-786">Name</span></span> | <span data-ttu-id="ee128-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-787">Type</span></span> | <span data-ttu-id="ee128-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-788">Attributes</span></span> | <span data-ttu-id="ee128-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ee128-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ee128-790">String &#124; Object</span></span>| | <span data-ttu-id="ee128-p147">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ee128-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ee128-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="ee128-793">**OR**</span></span><br/><span data-ttu-id="ee128-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ee128-796">String</span><span class="sxs-lookup"><span data-stu-id="ee128-796">String</span></span> | <span data-ttu-id="ee128-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-797">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-p149">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ee128-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ee128-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ee128-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-801">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-802">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="ee128-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ee128-803">String</span><span class="sxs-lookup"><span data-stu-id="ee128-803">String</span></span> | | <span data-ttu-id="ee128-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ee128-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ee128-806">String</span><span class="sxs-lookup"><span data-stu-id="ee128-806">String</span></span> | | <span data-ttu-id="ee128-807">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="ee128-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ee128-808">String</span><span class="sxs-lookup"><span data-stu-id="ee128-808">String</span></span> | | <span data-ttu-id="ee128-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ee128-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ee128-811">Booliano</span><span class="sxs-lookup"><span data-stu-id="ee128-811">Boolean</span></span> | | <span data-ttu-id="ee128-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ee128-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ee128-814">String</span><span class="sxs-lookup"><span data-stu-id="ee128-814">String</span></span> | | <span data-ttu-id="ee128-p153">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ee128-818">function</span><span class="sxs-lookup"><span data-stu-id="ee128-818">function</span></span> | <span data-ttu-id="ee128-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-819">&lt;optional&gt;</span></span> | <span data-ttu-id="ee128-820">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-821">Requirements</span></span>

|<span data-ttu-id="ee128-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-822">Requirement</span></span>| <span data-ttu-id="ee128-823">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-824">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-825">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-825">1.0</span></span>|
|[<span data-ttu-id="ee128-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-827">ReadItem</span></span>|
|[<span data-ttu-id="ee128-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-829">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ee128-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ee128-830">Examples</span></span>

<span data-ttu-id="ee128-831">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ee128-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ee128-832">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="ee128-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ee128-833">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="ee128-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ee128-834">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="ee128-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ee128-835">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ee128-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ee128-836">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="ee128-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ee128-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="ee128-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ee128-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-839">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-840">Requirements</span></span>

|<span data-ttu-id="ee128-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-841">Requirement</span></span>| <span data-ttu-id="ee128-842">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-843">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-844">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-844">1.0</span></span>|
|[<span data-ttu-id="ee128-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-846">ReadItem</span></span>|
|[<span data-ttu-id="ee128-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-848">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-849">Returns:</span></span>

<span data-ttu-id="ee128-850">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ee128-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ee128-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-851">Example</span></span>

<span data-ttu-id="ee128-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-852">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="ee128-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ee128-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ee128-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ee128-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-855">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-856">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-856">Parameters:</span></span>

|<span data-ttu-id="ee128-857">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-857">Name</span></span>| <span data-ttu-id="ee128-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-858">Type</span></span>| <span data-ttu-id="ee128-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ee128-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ee128-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="ee128-861">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="ee128-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-862">Requirements</span></span>

|<span data-ttu-id="ee128-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-863">Requirement</span></span>| <span data-ttu-id="ee128-864">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-865">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-866">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-866">1.0</span></span>|
|[<span data-ttu-id="ee128-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="ee128-868">Restricted</span></span>|
|[<span data-ttu-id="ee128-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-870">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-871">Returns:</span></span>

<span data-ttu-id="ee128-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="ee128-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ee128-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="ee128-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="ee128-874">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ee128-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ee128-875">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ee128-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="ee128-876">Value of `entityType`</span></span> | <span data-ttu-id="ee128-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="ee128-877">Type of objects in returned array</span></span> | <span data-ttu-id="ee128-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="ee128-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ee128-879">String</span><span class="sxs-lookup"><span data-stu-id="ee128-879">String</span></span> | <span data-ttu-id="ee128-880">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="ee128-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ee128-881">Contato</span><span class="sxs-lookup"><span data-stu-id="ee128-881">Contact</span></span> | <span data-ttu-id="ee128-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ee128-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ee128-883">String</span><span class="sxs-lookup"><span data-stu-id="ee128-883">String</span></span> | <span data-ttu-id="ee128-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ee128-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ee128-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ee128-885">MeetingSuggestion</span></span> | <span data-ttu-id="ee128-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ee128-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ee128-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ee128-887">PhoneNumber</span></span> | <span data-ttu-id="ee128-888">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="ee128-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ee128-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ee128-889">TaskSuggestion</span></span> | <span data-ttu-id="ee128-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ee128-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ee128-891">String</span><span class="sxs-lookup"><span data-stu-id="ee128-891">String</span></span> | <span data-ttu-id="ee128-892">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="ee128-892">**Restricted**</span></span> |

<span data-ttu-id="ee128-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ee128-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="ee128-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-894">Example</span></span>

<span data-ttu-id="ee128-895">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ee128-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="ee128-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ee128-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ee128-897">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ee128-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-898">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="ee128-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-900">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-900">Parameters:</span></span>

|<span data-ttu-id="ee128-901">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-901">Name</span></span>| <span data-ttu-id="ee128-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-902">Type</span></span>| <span data-ttu-id="ee128-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ee128-904">String</span><span class="sxs-lookup"><span data-stu-id="ee128-904">String</span></span>|<span data-ttu-id="ee128-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="ee128-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-906">Requirements</span></span>

|<span data-ttu-id="ee128-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-907">Requirement</span></span>| <span data-ttu-id="ee128-908">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-909">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-910">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-910">1.0</span></span>|
|[<span data-ttu-id="ee128-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-912">ReadItem</span></span>|
|[<span data-ttu-id="ee128-913">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-914">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-915">Returns:</span></span>

<span data-ttu-id="ee128-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="ee128-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ee128-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ee128-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="ee128-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ee128-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ee128-920">Retorna valores de cadeia de caracteres do item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ee128-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-921">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-p156">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="ee128-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ee128-925">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="ee128-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ee128-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ee128-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ee128-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="ee128-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-930">Requirements</span></span>

|<span data-ttu-id="ee128-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-931">Requirement</span></span>| <span data-ttu-id="ee128-932">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-933">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-934">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-934">1.0</span></span>|
|[<span data-ttu-id="ee128-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-936">ReadItem</span></span>|
|[<span data-ttu-id="ee128-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-938">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-939">Returns:</span></span>

<span data-ttu-id="ee128-p158">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="ee128-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ee128-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="ee128-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ee128-943">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ee128-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-944">Example</span></span>

<span data-ttu-id="ee128-945">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="ee128-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ee128-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ee128-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ee128-947">Retorna valores de cadeia de caracteres do item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ee128-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-948">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-949">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="ee128-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ee128-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="ee128-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-952">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-952">Parameters:</span></span>

|<span data-ttu-id="ee128-953">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-953">Name</span></span>| <span data-ttu-id="ee128-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-954">Type</span></span>| <span data-ttu-id="ee128-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ee128-956">String</span><span class="sxs-lookup"><span data-stu-id="ee128-956">String</span></span>|<span data-ttu-id="ee128-957">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="ee128-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-958">Requirements</span></span>

|<span data-ttu-id="ee128-959">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-959">Requirement</span></span>| <span data-ttu-id="ee128-960">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-961">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-962">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-962">1.0</span></span>|
|[<span data-ttu-id="ee128-963">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-964">ReadItem</span></span>|
|[<span data-ttu-id="ee128-965">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-966">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-967">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-967">Returns:</span></span>

<span data-ttu-id="ee128-968">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ee128-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ee128-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="ee128-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ee128-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ee128-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ee128-971">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-971">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ee128-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ee128-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ee128-973">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ee128-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="ee128-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-976">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-976">Parameters:</span></span>

|<span data-ttu-id="ee128-977">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-977">Name</span></span>| <span data-ttu-id="ee128-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-978">Type</span></span>| <span data-ttu-id="ee128-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-979">Attributes</span></span>| <span data-ttu-id="ee128-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="ee128-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ee128-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ee128-p161">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="ee128-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="ee128-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-985">Object</span></span>| <span data-ttu-id="ee128-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-986">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-987">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ee128-988">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-988">Object</span></span>| <span data-ttu-id="ee128-989">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-989">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-990">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ee128-991">function</span><span class="sxs-lookup"><span data-stu-id="ee128-991">function</span></span>||<span data-ttu-id="ee128-992">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ee128-993">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="ee128-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ee128-994">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="ee128-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-995">Requirements</span></span>

|<span data-ttu-id="ee128-996">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-996">Requirement</span></span>| <span data-ttu-id="ee128-997">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-998">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-999">1.2</span><span class="sxs-lookup"><span data-stu-id="ee128-999">1.2</span></span>|
|[<span data-ttu-id="ee128-1000">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-1002">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1003">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-1004">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-1004">Returns:</span></span>

<span data-ttu-id="ee128-1005">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="ee128-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="ee128-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="ee128-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ee128-1007">String</span><span class="sxs-lookup"><span data-stu-id="ee128-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ee128-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="ee128-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ee128-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="ee128-p163">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ee128-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-1012">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1013">Requirements</span></span>

|<span data-ttu-id="ee128-1014">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1014">Requirement</span></span>| <span data-ttu-id="ee128-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1016">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="ee128-1017">-16</span></span> |
|[<span data-ttu-id="ee128-1018">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1019">ReadItem</span></span>|
|[<span data-ttu-id="ee128-1020">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1021">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-1022">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-1022">Returns:</span></span>

<span data-ttu-id="ee128-1023">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ee128-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ee128-1024">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1024">Example</span></span>

<span data-ttu-id="ee128-1025">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="ee128-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ee128-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ee128-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ee128-p164">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ee128-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-1029">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ee128-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ee128-p165">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="ee128-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ee128-1033">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="ee128-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ee128-1034">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ee128-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ee128-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="ee128-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee128-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1038">Requirements</span></span>

|<span data-ttu-id="ee128-1039">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1039">Requirement</span></span>| <span data-ttu-id="ee128-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1041">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="ee128-1042">-16</span></span> |
|[<span data-ttu-id="ee128-1043">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1044">ReadItem</span></span>|
|[<span data-ttu-id="ee128-1045">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1046">Leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ee128-1047">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ee128-1047">Returns:</span></span>

<span data-ttu-id="ee128-p167">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="ee128-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ee128-1050">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1050">Example</span></span>

<span data-ttu-id="ee128-1051">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="ee128-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ee128-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ee128-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ee128-1053">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ee128-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ee128-p168">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="ee128-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-1057">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-1057">Parameters:</span></span>

|<span data-ttu-id="ee128-1058">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-1058">Name</span></span>| <span data-ttu-id="ee128-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-1059">Type</span></span>| <span data-ttu-id="ee128-1060">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-1060">Attributes</span></span>| <span data-ttu-id="ee128-1061">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ee128-1062">function</span><span class="sxs-lookup"><span data-stu-id="ee128-1062">function</span></span>||<span data-ttu-id="ee128-1063">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ee128-1064">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ee128-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ee128-1065">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="ee128-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ee128-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1066">Object</span></span>| <span data-ttu-id="ee128-1067">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1068">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="ee128-1069">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1070">Requirements</span></span>

|<span data-ttu-id="ee128-1071">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1071">Requirement</span></span>| <span data-ttu-id="ee128-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1073">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="ee128-1074">1.0</span></span>|
|[<span data-ttu-id="ee128-1075">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1076">ReadItem</span></span>|
|[<span data-ttu-id="ee128-1077">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1078">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee128-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1079">Example</span></span>

<span data-ttu-id="ee128-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ee128-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ee128-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ee128-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ee128-1084">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ee128-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ee128-p172">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="ee128-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-1089">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-1089">Parameters:</span></span>

|<span data-ttu-id="ee128-1090">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-1090">Name</span></span>| <span data-ttu-id="ee128-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-1091">Type</span></span>| <span data-ttu-id="ee128-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-1092">Attributes</span></span>| <span data-ttu-id="ee128-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ee128-1094">String</span><span class="sxs-lookup"><span data-stu-id="ee128-1094">String</span></span>||<span data-ttu-id="ee128-p173">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee128-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="ee128-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1097">Object</span></span>| <span data-ttu-id="ee128-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1099">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ee128-1100">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1100">Object</span></span>| <span data-ttu-id="ee128-1101">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1102">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ee128-1103">function</span><span class="sxs-lookup"><span data-stu-id="ee128-1103">function</span></span>| <span data-ttu-id="ee128-1104">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1105">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ee128-1106">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="ee128-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ee128-1107">Erros</span><span class="sxs-lookup"><span data-stu-id="ee128-1107">Errors</span></span>

| <span data-ttu-id="ee128-1108">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ee128-1108">Error code</span></span> | <span data-ttu-id="ee128-1109">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ee128-1110">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="ee128-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-1111">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1111">Requirements</span></span>

|<span data-ttu-id="ee128-1112">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1112">Requirement</span></span>| <span data-ttu-id="ee128-1113">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1114">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="ee128-1115">1.1</span></span>|
|[<span data-ttu-id="ee128-1116">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-1118">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1119">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-1120">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1120">Example</span></span>

<span data-ttu-id="ee128-1121">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="ee128-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="ee128-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ee128-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="ee128-1123">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ee128-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="ee128-p174">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="ee128-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-1127">Se seu suplemento chamar `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="ee128-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="ee128-1128">Até que o item esteja sincronizado, usar `itemId` retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="ee128-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ee128-p176">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="ee128-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ee128-1132">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="ee128-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ee128-1133">O Outlook para Mac não dá suporte a `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ee128-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="ee128-1134">Chamar `saveAsync` em uma reunião no Outlook para Mac retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="ee128-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="ee128-1135">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ee128-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-1136">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-1136">Parameters:</span></span>

|<span data-ttu-id="ee128-1137">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-1137">Name</span></span>| <span data-ttu-id="ee128-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-1138">Type</span></span>| <span data-ttu-id="ee128-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-1139">Attributes</span></span>| <span data-ttu-id="ee128-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="ee128-1141">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1141">Object</span></span>| <span data-ttu-id="ee128-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1143">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ee128-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1144">Object</span></span>| <span data-ttu-id="ee128-1145">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1146">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ee128-1147">function</span><span class="sxs-lookup"><span data-stu-id="ee128-1147">function</span></span>||<span data-ttu-id="ee128-1148">Quando o método é concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ee128-1149">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ee128-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee128-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1150">Requirements</span></span>

|<span data-ttu-id="ee128-1151">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1151">Requirement</span></span>| <span data-ttu-id="ee128-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1153">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="ee128-1154">1.3</span></span>|
|[<span data-ttu-id="ee128-1155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-1157">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1158">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ee128-1159">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ee128-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="ee128-p178">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="ee128-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ee128-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ee128-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ee128-1163">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ee128-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ee128-p179">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="ee128-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ee128-1167">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ee128-1167">Parameters:</span></span>

|<span data-ttu-id="ee128-1168">Nome</span><span class="sxs-lookup"><span data-stu-id="ee128-1168">Name</span></span>| <span data-ttu-id="ee128-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="ee128-1169">Type</span></span>| <span data-ttu-id="ee128-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee128-1170">Attributes</span></span>| <span data-ttu-id="ee128-1171">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee128-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ee128-1172">String</span><span class="sxs-lookup"><span data-stu-id="ee128-1172">String</span></span>||<span data-ttu-id="ee128-p180">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="ee128-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="ee128-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1176">Object</span></span>| <span data-ttu-id="ee128-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1178">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee128-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ee128-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="ee128-1179">Object</span></span>| <span data-ttu-id="ee128-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-1181">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ee128-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="ee128-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ee128-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="ee128-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ee128-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="ee128-p181">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="ee128-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ee128-p182">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="ee128-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ee128-1188">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="ee128-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="ee128-1189">function</span><span class="sxs-lookup"><span data-stu-id="ee128-1189">function</span></span>||<span data-ttu-id="ee128-1190">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ee128-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee128-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee128-1191">Requirements</span></span>

|<span data-ttu-id="ee128-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee128-1192">Requirement</span></span>| <span data-ttu-id="ee128-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="ee128-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee128-1194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee128-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee128-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="ee128-1195">1.2</span></span>|
|[<span data-ttu-id="ee128-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee128-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee128-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ee128-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="ee128-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee128-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee128-1199">Redação</span><span class="sxs-lookup"><span data-stu-id="ee128-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ee128-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee128-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```