
# <a name="item"></a><span data-ttu-id="ab066-101">item</span><span class="sxs-lookup"><span data-stu-id="ab066-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ab066-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ab066-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ab066-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ab066-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-105">Requirements</span></span>

|<span data-ttu-id="ab066-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-106">Requirement</span></span>| <span data-ttu-id="ab066-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-109">1.0</span></span>|
|[<span data-ttu-id="ab066-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="ab066-111">Restricted</span></span>|
|[<span data-ttu-id="ab066-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ab066-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ab066-114">Members and methods</span></span>

| <span data-ttu-id="ab066-115">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-115">Member</span></span> | <span data-ttu-id="ab066-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ab066-117">attachments</span><span class="sxs-lookup"><span data-stu-id="ab066-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="ab066-118">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-118">Member</span></span> |
| [<span data-ttu-id="ab066-119">bcc</span><span class="sxs-lookup"><span data-stu-id="ab066-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ab066-120">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-120">Member</span></span> |
| [<span data-ttu-id="ab066-121">body</span><span class="sxs-lookup"><span data-stu-id="ab066-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="ab066-122">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-122">Member</span></span> |
| [<span data-ttu-id="ab066-123">cc</span><span class="sxs-lookup"><span data-stu-id="ab066-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ab066-124">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-124">Member</span></span> |
| [<span data-ttu-id="ab066-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="ab066-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ab066-126">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-126">Member</span></span> |
| [<span data-ttu-id="ab066-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ab066-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ab066-128">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-128">Member</span></span> |
| [<span data-ttu-id="ab066-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ab066-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ab066-130">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-130">Member</span></span> |
| [<span data-ttu-id="ab066-131">end</span><span class="sxs-lookup"><span data-stu-id="ab066-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="ab066-132">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-132">Member</span></span> |
| [<span data-ttu-id="ab066-133">from</span><span class="sxs-lookup"><span data-stu-id="ab066-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ab066-134">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-134">Member</span></span> |
| [<span data-ttu-id="ab066-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ab066-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ab066-136">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-136">Member</span></span> |
| [<span data-ttu-id="ab066-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="ab066-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ab066-138">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-138">Member</span></span> |
| [<span data-ttu-id="ab066-139">itemId</span><span class="sxs-lookup"><span data-stu-id="ab066-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ab066-140">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-140">Member</span></span> |
| [<span data-ttu-id="ab066-141">itemType</span><span class="sxs-lookup"><span data-stu-id="ab066-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="ab066-142">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-142">Member</span></span> |
| [<span data-ttu-id="ab066-143">location</span><span class="sxs-lookup"><span data-stu-id="ab066-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="ab066-144">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-144">Member</span></span> |
| [<span data-ttu-id="ab066-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ab066-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ab066-146">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-146">Member</span></span> |
| [<span data-ttu-id="ab066-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ab066-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="ab066-148">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-148">Member</span></span> |
| [<span data-ttu-id="ab066-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ab066-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ab066-150">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-150">Member</span></span> |
| [<span data-ttu-id="ab066-151">organizer</span><span class="sxs-lookup"><span data-stu-id="ab066-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ab066-152">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-152">Member</span></span> |
| [<span data-ttu-id="ab066-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ab066-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ab066-154">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-154">Member</span></span> |
| [<span data-ttu-id="ab066-155">sender</span><span class="sxs-lookup"><span data-stu-id="ab066-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="ab066-156">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-156">Member</span></span> |
| [<span data-ttu-id="ab066-157">start</span><span class="sxs-lookup"><span data-stu-id="ab066-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="ab066-158">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-158">Member</span></span> |
| [<span data-ttu-id="ab066-159">subject</span><span class="sxs-lookup"><span data-stu-id="ab066-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="ab066-160">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-160">Member</span></span> |
| [<span data-ttu-id="ab066-161">to</span><span class="sxs-lookup"><span data-stu-id="ab066-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="ab066-162">Membro</span><span class="sxs-lookup"><span data-stu-id="ab066-162">Member</span></span> |
| [<span data-ttu-id="ab066-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ab066-164">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-164">Method</span></span> |
| [<span data-ttu-id="ab066-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ab066-166">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-166">Method</span></span> |
| [<span data-ttu-id="ab066-167">close</span><span class="sxs-lookup"><span data-stu-id="ab066-167">close</span></span>](#close) | <span data-ttu-id="ab066-168">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-168">Method</span></span> |
| [<span data-ttu-id="ab066-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ab066-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="ab066-170">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-170">Method</span></span> |
| [<span data-ttu-id="ab066-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ab066-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="ab066-172">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-172">Method</span></span> |
| [<span data-ttu-id="ab066-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="ab066-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="ab066-174">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-174">Method</span></span> |
| [<span data-ttu-id="ab066-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ab066-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="ab066-176">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-176">Method</span></span> |
| [<span data-ttu-id="ab066-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ab066-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="ab066-178">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-178">Method</span></span> |
| [<span data-ttu-id="ab066-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ab066-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ab066-180">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-180">Method</span></span> |
| [<span data-ttu-id="ab066-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ab066-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ab066-182">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-182">Method</span></span> |
| [<span data-ttu-id="ab066-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ab066-184">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-184">Method</span></span> |
| [<span data-ttu-id="ab066-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="ab066-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="ab066-186">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-186">Method</span></span> |
| [<span data-ttu-id="ab066-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ab066-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="ab066-188">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-188">Method</span></span> |
| [<span data-ttu-id="ab066-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ab066-190">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-190">Method</span></span> |
| [<span data-ttu-id="ab066-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ab066-192">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-192">Method</span></span> |
| [<span data-ttu-id="ab066-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ab066-194">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-194">Method</span></span> |
| [<span data-ttu-id="ab066-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ab066-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ab066-196">Método</span><span class="sxs-lookup"><span data-stu-id="ab066-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ab066-197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-197">Example</span></span>

<span data-ttu-id="ab066-198">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ab066-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="ab066-199">Membros</span><span class="sxs-lookup"><span data-stu-id="ab066-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="ab066-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ab066-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="ab066-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-203">Certos tipos de arquivos são bloqueados pelo Outlook devido a potenciais problemas de segurança e portanto não são retornados.</span><span class="sxs-lookup"><span data-stu-id="ab066-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ab066-204">Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ab066-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-205">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-205">Type:</span></span>

*   <span data-ttu-id="ab066-206">Array. <[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ab066-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-207">Requirements</span></span>

|<span data-ttu-id="ab066-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-208">Requirement</span></span>| <span data-ttu-id="ab066-209">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-210">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-211">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-211">1.0</span></span>|
|[<span data-ttu-id="ab066-212">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-213">ReadItem</span></span>|
|[<span data-ttu-id="ab066-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-215">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-216">Example</span></span>

<span data-ttu-id="ab066-217">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ab066-218">cco:[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ab066-219">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ab066-220">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ab066-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-221">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-221">Type:</span></span>

*   [<span data-ttu-id="ab066-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="ab066-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ab066-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-223">Requirements</span></span>

|<span data-ttu-id="ab066-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-224">Requirement</span></span>| <span data-ttu-id="ab066-225">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-226">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-227">1.1</span><span class="sxs-lookup"><span data-stu-id="ab066-227">1.1</span></span>|
|[<span data-ttu-id="ab066-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-229">ReadItem</span></span>|
|[<span data-ttu-id="ab066-230">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-231">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-232">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="ab066-233">corpo:[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="ab066-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="ab066-234">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="ab066-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-235">Type:</span></span>

*   [<span data-ttu-id="ab066-236">Body</span><span class="sxs-lookup"><span data-stu-id="ab066-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="ab066-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-237">Requirements</span></span>

|<span data-ttu-id="ab066-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-238">Requirement</span></span>| <span data-ttu-id="ab066-239">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-240">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-241">1.1</span><span class="sxs-lookup"><span data-stu-id="ab066-241">1.1</span></span>|
|[<span data-ttu-id="ab066-242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-243">ReadItem</span></span>|
|[<span data-ttu-id="ab066-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-245">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ab066-246">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ab066-247">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ab066-248">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-249">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-249">Read mode</span></span>

<span data-ttu-id="ab066-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="ab066-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-252">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-252">Compose mode</span></span>

<span data-ttu-id="ab066-253">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-254">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-254">Type:</span></span>

*   <span data-ttu-id="ab066-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-256">Requirements</span></span>

|<span data-ttu-id="ab066-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-257">Requirement</span></span>| <span data-ttu-id="ab066-258">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-259">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-260">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-260">1.0</span></span>|
|[<span data-ttu-id="ab066-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-262">ReadItem</span></span>|
|[<span data-ttu-id="ab066-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-264">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="ab066-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="ab066-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="ab066-267">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="ab066-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ab066-p107">Se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação, você obterá um número inteiro nessa propriedade. Se posteriormente o usuário alterar o assunto da mensagem ao enviar a resposta, o ID da conversa daquela mensagem será alterado e o valor obtido anteriormente não se aplicará mais.</span><span class="sxs-lookup"><span data-stu-id="ab066-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ab066-p108">Para um novo item em um formulário de redação, o valor dessa propriedade é nulo. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="ab066-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-272">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-272">Type:</span></span>

*   <span data-ttu-id="ab066-273">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-274">Requirements</span></span>

|<span data-ttu-id="ab066-275">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-275">Requirement</span></span>| <span data-ttu-id="ab066-276">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-277">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-278">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-278">1.0</span></span>|
|[<span data-ttu-id="ab066-279">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-280">ReadItem</span></span>|
|[<span data-ttu-id="ab066-281">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-282">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="ab066-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="ab066-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="ab066-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-286">Type:</span></span>

*   <span data-ttu-id="ab066-287">Data</span><span class="sxs-lookup"><span data-stu-id="ab066-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-288">Requirements</span></span>

|<span data-ttu-id="ab066-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-289">Requirement</span></span>| <span data-ttu-id="ab066-290">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-291">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-292">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-292">1.0</span></span>|
|[<span data-ttu-id="ab066-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-294">ReadItem</span></span>|
|[<span data-ttu-id="ab066-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-296">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ab066-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="ab066-298">dateTimeModified :Date</span></span>

<span data-ttu-id="ab066-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-301">Esse membro não é compatível com Outlook para iOS ou Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-302">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-302">Type:</span></span>

*   <span data-ttu-id="ab066-303">Data</span><span class="sxs-lookup"><span data-stu-id="ab066-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-304">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-304">Requirements</span></span>

|<span data-ttu-id="ab066-305">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-305">Requirement</span></span>| <span data-ttu-id="ab066-306">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-307">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-308">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-308">1.0</span></span>|
|[<span data-ttu-id="ab066-309">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-310">ReadItem</span></span>|
|[<span data-ttu-id="ab066-311">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-312">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="ab066-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ab066-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="ab066-315">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="ab066-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ab066-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor da propriedade para a data e hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="ab066-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-318">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-318">Read mode</span></span>

<span data-ttu-id="ab066-319">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ab066-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-320">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-320">Compose mode</span></span>

<span data-ttu-id="ab066-321">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ab066-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ab066-322">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)  para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)  para converter a hora local no cliente para UTC no servidor.</span><span class="sxs-lookup"><span data-stu-id="ab066-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-323">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-323">Type:</span></span>

*   <span data-ttu-id="ab066-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ab066-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-325">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-325">Requirements</span></span>

|<span data-ttu-id="ab066-326">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-326">Requirement</span></span>| <span data-ttu-id="ab066-327">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-328">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-329">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-329">1.0</span></span>|
|[<span data-ttu-id="ab066-330">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-331">ReadItem</span></span>|
|[<span data-ttu-id="ab066-332">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-333">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-334">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-334">Example</span></span>

<span data-ttu-id="ab066-335">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ab066-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ab066-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ab066-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ab066-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ab066-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegador, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="ab066-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-341">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ab066-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-342">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-342">Type:</span></span>

*   [<span data-ttu-id="ab066-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ab066-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ab066-344">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-344">Requirements</span></span>

|<span data-ttu-id="ab066-345">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-345">Requirement</span></span>| <span data-ttu-id="ab066-346">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-347">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-348">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-348">1.0</span></span>|
|[<span data-ttu-id="ab066-349">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-350">ReadItem</span></span>|
|[<span data-ttu-id="ab066-351">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-352">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="ab066-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="ab066-353">internetMessageId :String</span></span>

<span data-ttu-id="ab066-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-356">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-356">Type:</span></span>

*   <span data-ttu-id="ab066-357">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-358">Requirements</span></span>

|<span data-ttu-id="ab066-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-359">Requirement</span></span>| <span data-ttu-id="ab066-360">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-361">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-362">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-362">1.0</span></span>|
|[<span data-ttu-id="ab066-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-364">ReadItem</span></span>|
|[<span data-ttu-id="ab066-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-366">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ab066-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="ab066-368">itemClass :String</span></span>

<span data-ttu-id="ab066-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ab066-p116">A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir estão as classes de mensagem padrão para itens de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ab066-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-373">Type</span></span> | <span data-ttu-id="ab066-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-374">Description</span></span> | <span data-ttu-id="ab066-375">classe do item</span><span class="sxs-lookup"><span data-stu-id="ab066-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ab066-376">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="ab066-376">Appointment items</span></span> | <span data-ttu-id="ab066-377">São itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="ab066-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="ab066-378">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="ab066-378">Message items</span></span> | <span data-ttu-id="ab066-379">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagem base.</span><span class="sxs-lookup"><span data-stu-id="ab066-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ab066-380">Você pode criar classes de mensagens personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="ab066-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-381">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-381">Type:</span></span>

*   <span data-ttu-id="ab066-382">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-383">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-383">Requirements</span></span>

|<span data-ttu-id="ab066-384">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-384">Requirement</span></span>| <span data-ttu-id="ab066-385">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-386">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-387">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-387">1.0</span></span>|
|[<span data-ttu-id="ab066-388">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-389">ReadItem</span></span>|
|[<span data-ttu-id="ab066-390">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-391">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-392">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ab066-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="ab066-393">(nullable) itemId :String</span></span>

<span data-ttu-id="ab066-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-396">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="ab066-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ab066-397">A propriedade `itemId` não é idêntica ao Entry ID do Outlook ou ao ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ab066-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ab066-398">Antes de fazer chamadas à API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="ab066-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ab066-399">Para mais informações, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ab066-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ab066-p119">A propriedade `itemId` não está disponível no modo de redação. Se  um identificador de item for requerido, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no repositório, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-402">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-402">Type:</span></span>

*   <span data-ttu-id="ab066-403">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-404">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-404">Requirements</span></span>

|<span data-ttu-id="ab066-405">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-405">Requirement</span></span>| <span data-ttu-id="ab066-406">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-407">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-408">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-408">1.0</span></span>|
|[<span data-ttu-id="ab066-409">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-410">ReadItem</span></span>|
|[<span data-ttu-id="ab066-411">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-412">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-413">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-413">Example</span></span>

<span data-ttu-id="ab066-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retornar `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item a partir do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="ab066-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="ab066-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ab066-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ab066-417">Obtém o tipo de item que uma instância representa.</span><span class="sxs-lookup"><span data-stu-id="ab066-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ab066-418">A propriedade `itemType` retorna um dos valores da enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-419">Type:</span></span>

*   [<span data-ttu-id="ab066-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ab066-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ab066-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-421">Requirements</span></span>

|<span data-ttu-id="ab066-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-422">Requirement</span></span>| <span data-ttu-id="ab066-423">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-424">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-425">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-425">1.0</span></span>|
|[<span data-ttu-id="ab066-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-427">ReadItem</span></span>|
|[<span data-ttu-id="ab066-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-429">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="ab066-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="ab066-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="ab066-432">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-433">Read mode</span></span>

<span data-ttu-id="ab066-434">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-435">Compose mode</span></span>

<span data-ttu-id="ab066-436">A propriedade `location` retorna um objeto `Location` que fornece métodos para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-437">Type:</span></span>

*   <span data-ttu-id="ab066-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="ab066-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-439">Requirements</span></span>

|<span data-ttu-id="ab066-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-440">Requirement</span></span>| <span data-ttu-id="ab066-441">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-442">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-443">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-443">1.0</span></span>|
|[<span data-ttu-id="ab066-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-445">ReadItem</span></span>|
|[<span data-ttu-id="ab066-446">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-447">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ab066-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="ab066-449">normalizedSubject :String</span></span>

<span data-ttu-id="ab066-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ab066-p122">A propriedade normalizedSubject obtém o assunto do item sem nenhum dos prefixos padrão (como `RE:` e `FW:`) que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="ab066-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-454">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-454">Type:</span></span>

*   <span data-ttu-id="ab066-455">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-456">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-456">Requirements</span></span>

|<span data-ttu-id="ab066-457">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-457">Requirement</span></span>| <span data-ttu-id="ab066-458">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-459">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-460">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-460">1.0</span></span>|
|[<span data-ttu-id="ab066-461">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-462">ReadItem</span></span>|
|[<span data-ttu-id="ab066-463">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-464">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-465">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="ab066-466">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="ab066-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="ab066-467">Obtém as mensagens de notificação para um item.</span><span class="sxs-lookup"><span data-stu-id="ab066-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-468">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-468">Type:</span></span>

*   [<span data-ttu-id="ab066-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ab066-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="ab066-470">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-470">Requirements</span></span>

|<span data-ttu-id="ab066-471">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-471">Requirement</span></span>| <span data-ttu-id="ab066-472">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-473">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-474">1.3</span><span class="sxs-lookup"><span data-stu-id="ab066-474">1.3</span></span>|
|[<span data-ttu-id="ab066-475">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-476">ReadItem</span></span>|
|[<span data-ttu-id="ab066-477">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-478">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ab066-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ab066-480">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="ab066-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ab066-481">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-482">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-482">Read mode</span></span>

<span data-ttu-id="ab066-483">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="ab066-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-484">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-484">Compose mode</span></span>

<span data-ttu-id="ab066-485">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="ab066-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-486">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-486">Type:</span></span>

*   <span data-ttu-id="ab066-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-488">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-488">Requirements</span></span>

|<span data-ttu-id="ab066-489">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-489">Requirement</span></span>| <span data-ttu-id="ab066-490">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-491">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-492">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-492">1.0</span></span>|
|[<span data-ttu-id="ab066-493">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-494">ReadItem</span></span>|
|[<span data-ttu-id="ab066-495">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-496">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-497">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ab066-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ab066-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ab066-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-501">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-501">Type:</span></span>

*   [<span data-ttu-id="ab066-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ab066-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ab066-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-503">Requirements</span></span>

|<span data-ttu-id="ab066-504">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-504">Requirement</span></span>| <span data-ttu-id="ab066-505">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-506">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-507">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-507">1.0</span></span>|
|[<span data-ttu-id="ab066-508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-509">ReadItem</span></span>|
|[<span data-ttu-id="ab066-510">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-511">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-512">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ab066-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ab066-514">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="ab066-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ab066-515">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-516">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-516">Read mode</span></span>

<span data-ttu-id="ab066-517">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="ab066-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-518">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-518">Compose mode</span></span>

<span data-ttu-id="ab066-519">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="ab066-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-520">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-520">Type:</span></span>

*   <span data-ttu-id="ab066-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-522">Requirements</span></span>

|<span data-ttu-id="ab066-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-523">Requirement</span></span>| <span data-ttu-id="ab066-524">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-525">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-526">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-526">1.0</span></span>|
|[<span data-ttu-id="ab066-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-528">ReadItem</span></span>|
|[<span data-ttu-id="ab066-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-530">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-531">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="ab066-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ab066-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="ab066-p126">Obtém o endereço de e-mail do remetente de uma mensagem de e-mail. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="ab066-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ab066-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegador, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="ab066-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-537">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ab066-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-538">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-538">Type:</span></span>

*   [<span data-ttu-id="ab066-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ab066-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ab066-540">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-540">Requirements</span></span>

|<span data-ttu-id="ab066-541">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-541">Requirement</span></span>| <span data-ttu-id="ab066-542">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-543">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-544">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-544">1.0</span></span>|
|[<span data-ttu-id="ab066-545">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-546">ReadItem</span></span>|
|[<span data-ttu-id="ab066-547">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-548">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-549">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="ab066-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ab066-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="ab066-551">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="ab066-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ab066-p128">A propriedade `start` é expressa como um valor de data e valor temporal no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="ab066-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-554">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-554">Read mode</span></span>

<span data-ttu-id="ab066-555">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="ab066-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-556">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-556">Compose mode</span></span>

<span data-ttu-id="ab066-557">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ab066-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ab066-558">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC no servidor.</span><span class="sxs-lookup"><span data-stu-id="ab066-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-559">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-559">Type:</span></span>

*   <span data-ttu-id="ab066-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="ab066-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-561">Requirements</span></span>

|<span data-ttu-id="ab066-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-562">Requirement</span></span>| <span data-ttu-id="ab066-563">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-564">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-565">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-565">1.0</span></span>|
|[<span data-ttu-id="ab066-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-567">ReadItem</span></span>|
|[<span data-ttu-id="ab066-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-569">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-570">Example</span></span>

<span data-ttu-id="ab066-571">O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="ab066-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="ab066-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ab066-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="ab066-573">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="ab066-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ab066-574">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="ab066-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-575">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-575">Read mode</span></span>

<span data-ttu-id="ab066-p129">A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ab066-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="ab066-578">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-578">Compose mode</span></span>

<span data-ttu-id="ab066-579">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="ab066-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ab066-580">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-580">Type:</span></span>

*   <span data-ttu-id="ab066-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ab066-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-582">Requirements</span></span>

|<span data-ttu-id="ab066-583">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-583">Requirement</span></span>| <span data-ttu-id="ab066-584">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-585">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-586">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-586">1.0</span></span>|
|[<span data-ttu-id="ab066-587">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-588">ReadItem</span></span>|
|[<span data-ttu-id="ab066-589">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-590">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="ab066-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="ab066-592">Fornece acesso aos destinatários na linha **To** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ab066-593">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ab066-594">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-594">Read mode</span></span>

<span data-ttu-id="ab066-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **To** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="ab066-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ab066-597">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="ab066-597">Compose mode</span></span>

<span data-ttu-id="ab066-598">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **To** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ab066-599">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ab066-599">Type:</span></span>

*   <span data-ttu-id="ab066-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ab066-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-601">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-601">Requirements</span></span>

|<span data-ttu-id="ab066-602">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-602">Requirement</span></span>| <span data-ttu-id="ab066-603">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-604">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-605">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-605">1.0</span></span>|
|[<span data-ttu-id="ab066-606">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-607">ReadItem</span></span>|
|[<span data-ttu-id="ab066-608">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-609">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-610">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="ab066-611">Métodos</span><span class="sxs-lookup"><span data-stu-id="ab066-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ab066-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ab066-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ab066-613">Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.</span><span class="sxs-lookup"><span data-stu-id="ab066-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ab066-614">O método `addFileAttachmentAsync` carrega o arquivo da URI especificada e o anexa ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="ab066-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ab066-615">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="ab066-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-616">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-616">Parameters:</span></span>

|<span data-ttu-id="ab066-617">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-617">Name</span></span>| <span data-ttu-id="ab066-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-618">Type</span></span>| <span data-ttu-id="ab066-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-619">Attributes</span></span>| <span data-ttu-id="ab066-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ab066-621">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-621">String</span></span>||<span data-ttu-id="ab066-p132">A URI que fornece a localização do arquivo a ser anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ab066-624">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-624">String</span></span>||<span data-ttu-id="ab066-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ab066-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-627">Object</span></span>| <span data-ttu-id="ab066-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-628">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-629">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="ab066-630">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-630">Object</span></span> | <span data-ttu-id="ab066-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-631">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-632">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="ab066-633">Booleano</span><span class="sxs-lookup"><span data-stu-id="ab066-633">Boolean</span></span> | <span data-ttu-id="ab066-634">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-634">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-635">Se for `true`, indicará que o anexo será embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ab066-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="ab066-636">função</span><span class="sxs-lookup"><span data-stu-id="ab066-636">function</span></span>| <span data-ttu-id="ab066-637">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-637">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-638">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ab066-639">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ab066-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ab066-640">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="ab066-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ab066-641">Erros</span><span class="sxs-lookup"><span data-stu-id="ab066-641">Errors</span></span>

| <span data-ttu-id="ab066-642">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ab066-642">Error code</span></span> | <span data-ttu-id="ab066-643">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ab066-644">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="ab066-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ab066-645">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="ab066-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ab066-646">A mensagem ou o compromisso tem anexos demais.</span><span class="sxs-lookup"><span data-stu-id="ab066-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-647">Requirements</span></span>

|<span data-ttu-id="ab066-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-648">Requirement</span></span>| <span data-ttu-id="ab066-649">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-650">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-651">1.1</span><span class="sxs-lookup"><span data-stu-id="ab066-651">1.1</span></span>|
|[<span data-ttu-id="ab066-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-654">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-655">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ab066-656">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ab066-656">Examples</span></span>

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

<span data-ttu-id="ab066-657">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ab066-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ab066-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ab066-659">Adiciona um item do Exchange, como uma mensagem, como um anexo à mensagem ou ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ab066-p134">O método `addItemAttachmentAsync` anexa o item com o identificador especificado do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="ab066-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ab066-663">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="ab066-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ab066-664">Se o suplemento do Office estiver sendo executado no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a outros itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="ab066-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-665">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-665">Parameters:</span></span>

|<span data-ttu-id="ab066-666">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-666">Name</span></span>| <span data-ttu-id="ab066-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-667">Type</span></span>| <span data-ttu-id="ab066-668">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-668">Attributes</span></span>| <span data-ttu-id="ab066-669">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ab066-670">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-670">String</span></span>||<span data-ttu-id="ab066-p135">O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ab066-673">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-673">String</span></span>||<span data-ttu-id="ab066-p136">O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ab066-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-676">Object</span></span>| <span data-ttu-id="ab066-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-677">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-678">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ab066-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-679">Object</span></span>| <span data-ttu-id="ab066-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-680">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ab066-682">função</span><span class="sxs-lookup"><span data-stu-id="ab066-682">function</span></span>| <span data-ttu-id="ab066-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-683">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-684">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ab066-685">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ab066-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ab066-686">Se não for possível adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` com a descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="ab066-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ab066-687">Erros</span><span class="sxs-lookup"><span data-stu-id="ab066-687">Errors</span></span>

| <span data-ttu-id="ab066-688">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ab066-688">Error code</span></span> | <span data-ttu-id="ab066-689">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ab066-690">A mensagem ou o compromisso tem anexos demais.</span><span class="sxs-lookup"><span data-stu-id="ab066-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-691">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-691">Requirements</span></span>

|<span data-ttu-id="ab066-692">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-692">Requirement</span></span>| <span data-ttu-id="ab066-693">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-694">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-695">1.1</span><span class="sxs-lookup"><span data-stu-id="ab066-695">1.1</span></span>|
|[<span data-ttu-id="ab066-696">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-698">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-699">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-700">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-700">Example</span></span>

<span data-ttu-id="ab066-701">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ab066-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="ab066-702">close()</span><span class="sxs-lookup"><span data-stu-id="ab066-702">close()</span></span>

<span data-ttu-id="ab066-703">Fecha o item atual que está sendo redigido.</span><span class="sxs-lookup"><span data-stu-id="ab066-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ab066-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item possuir alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação de fechamento.</span><span class="sxs-lookup"><span data-stu-id="ab066-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-706">No Outlook na Web, se o item for um compromisso e tiver sido salvo anteriormente usando `saveAsync`, será solicitado ao usuário para salvar, descartar ou cancelar, mesmo que nenhuma alteração tenha ocorrido após o item ter sido salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="ab066-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ab066-707">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="ab066-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-708">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-708">Requirements</span></span>

|<span data-ttu-id="ab066-709">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-709">Requirement</span></span>| <span data-ttu-id="ab066-710">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-711">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-712">1.3</span><span class="sxs-lookup"><span data-stu-id="ab066-712">1.3</span></span>|
|[<span data-ttu-id="ab066-713">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-714">Restrito</span><span class="sxs-lookup"><span data-stu-id="ab066-714">Restricted</span></span>|
|[<span data-ttu-id="ab066-715">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-716">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="ab066-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ab066-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="ab066-718">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="ab066-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-719">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-720">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="ab066-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ab066-721">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="ab066-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ab066-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="ab066-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-725">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-725">Parameters:</span></span>

| <span data-ttu-id="ab066-726">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-726">Name</span></span> | <span data-ttu-id="ab066-727">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-727">Type</span></span> | <span data-ttu-id="ab066-728">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-728">Attributes</span></span> | <span data-ttu-id="ab066-729">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ab066-730">String | Object</span><span class="sxs-lookup"><span data-stu-id="ab066-730">String &#124; Object</span></span>| |<span data-ttu-id="ab066-p139">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ab066-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ab066-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="ab066-733">**OR**</span></span><br/><span data-ttu-id="ab066-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="ab066-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ab066-736">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-736">String</span></span> | <span data-ttu-id="ab066-737">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-737">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-p141">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ab066-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ab066-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ab066-741">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-741">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-742">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ab066-743">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-743">String</span></span> | | <span data-ttu-id="ab066-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ab066-746">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-746">String</span></span> | | <span data-ttu-id="ab066-747">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="ab066-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ab066-748">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-748">String</span></span> | | <span data-ttu-id="ab066-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ab066-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ab066-751">Booleano</span><span class="sxs-lookup"><span data-stu-id="ab066-751">Boolean</span></span> | | <span data-ttu-id="ab066-p144">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ab066-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ab066-754">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-754">String</span></span> | | <span data-ttu-id="ab066-p145">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ab066-758">função</span><span class="sxs-lookup"><span data-stu-id="ab066-758">function</span></span> | <span data-ttu-id="ab066-759">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-759">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-760">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-761">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-761">Requirements</span></span>

|<span data-ttu-id="ab066-762">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-762">Requirement</span></span>| <span data-ttu-id="ab066-763">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-764">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-765">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-765">1.0</span></span>|
|[<span data-ttu-id="ab066-766">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-767">ReadItem</span></span>|
|[<span data-ttu-id="ab066-768">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-769">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ab066-770">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ab066-770">Examples</span></span>

<span data-ttu-id="ab066-771">O código a seguir passa uma sequência de caracteres para a função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ab066-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ab066-772">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="ab066-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ab066-773">Resposta apenas com o corpo.</span><span class="sxs-lookup"><span data-stu-id="ab066-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ab066-774">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="ab066-774">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="ab066-775">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-775">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="ab066-776">Resposta com o corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="ab066-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ab066-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="ab066-778">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="ab066-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-779">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-780">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="ab066-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ab066-781">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="ab066-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ab066-p146">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="ab066-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-785">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-785">Parameters:</span></span>

| <span data-ttu-id="ab066-786">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-786">Name</span></span> | <span data-ttu-id="ab066-787">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-787">Type</span></span> | <span data-ttu-id="ab066-788">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-788">Attributes</span></span> | <span data-ttu-id="ab066-789">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ab066-790">String | Object</span><span class="sxs-lookup"><span data-stu-id="ab066-790">String &#124; Object</span></span>| | <span data-ttu-id="ab066-p147">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ab066-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ab066-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="ab066-793">**OR**</span></span><br/><span data-ttu-id="ab066-p148">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="ab066-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ab066-796">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-796">String</span></span> | <span data-ttu-id="ab066-797">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-797">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-p149">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ab066-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ab066-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ab066-801">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-801">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-802">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ab066-803">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-803">String</span></span> | | <span data-ttu-id="ab066-p150">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ab066-806">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-806">String</span></span> | | <span data-ttu-id="ab066-807">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="ab066-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ab066-808">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-808">String</span></span> | | <span data-ttu-id="ab066-p151">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ab066-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ab066-811">Booleano</span><span class="sxs-lookup"><span data-stu-id="ab066-811">Boolean</span></span> | | <span data-ttu-id="ab066-p152">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="ab066-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ab066-814">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-814">String</span></span> | | <span data-ttu-id="ab066-p153">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ab066-818">função</span><span class="sxs-lookup"><span data-stu-id="ab066-818">function</span></span> | <span data-ttu-id="ab066-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-819">&lt;optional&gt;</span></span> | <span data-ttu-id="ab066-820">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-821">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-821">Requirements</span></span>

|<span data-ttu-id="ab066-822">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-822">Requirement</span></span>| <span data-ttu-id="ab066-823">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-824">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-825">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-825">1.0</span></span>|
|[<span data-ttu-id="ab066-826">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-827">ReadItem</span></span>|
|[<span data-ttu-id="ab066-828">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-829">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ab066-830">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ab066-830">Examples</span></span>

<span data-ttu-id="ab066-831">O código a seguir passa uma sequência de caracteres para a função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ab066-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ab066-832">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="ab066-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ab066-833">Resposta apenas com o corpo.</span><span class="sxs-lookup"><span data-stu-id="ab066-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ab066-834">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="ab066-834">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="ab066-835">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="ab066-835">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="ab066-836">Resposta com o corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="ab066-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ab066-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="ab066-838">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ab066-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-839">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-840">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-840">Requirements</span></span>

|<span data-ttu-id="ab066-841">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-841">Requirement</span></span>| <span data-ttu-id="ab066-842">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-843">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-844">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-844">1.0</span></span>|
|[<span data-ttu-id="ab066-845">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-846">ReadItem</span></span>|
|[<span data-ttu-id="ab066-847">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-848">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-849">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-849">Returns:</span></span>

<span data-ttu-id="ab066-850">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ab066-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ab066-851">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-851">Example</span></span>

<span data-ttu-id="ab066-852">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="ab066-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ab066-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ab066-854">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ab066-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-855">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-856">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-856">Parameters:</span></span>

|<span data-ttu-id="ab066-857">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-857">Name</span></span>| <span data-ttu-id="ab066-858">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-858">Type</span></span>| <span data-ttu-id="ab066-859">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ab066-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ab066-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="ab066-861">Um dos valores da enumeração EntityType.</span><span class="sxs-lookup"><span data-stu-id="ab066-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-862">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-862">Requirements</span></span>

|<span data-ttu-id="ab066-863">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-863">Requirement</span></span>| <span data-ttu-id="ab066-864">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-865">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-866">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-866">1.0</span></span>|
|[<span data-ttu-id="ab066-867">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-868">Restrito</span><span class="sxs-lookup"><span data-stu-id="ab066-868">Restricted</span></span>|
|[<span data-ttu-id="ab066-869">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-870">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-871">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-871">Returns:</span></span>

<span data-ttu-id="ab066-872">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="ab066-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ab066-873">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="ab066-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="ab066-874">Caso contrário, o tipo dos objetos na matriz retornada dependerá do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ab066-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ab066-875">Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ab066-876">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="ab066-876">Value of `entityType`</span></span> | <span data-ttu-id="ab066-877">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="ab066-877">Type of objects in returned array</span></span> | <span data-ttu-id="ab066-878">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="ab066-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ab066-879">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-879">String</span></span> | <span data-ttu-id="ab066-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ab066-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ab066-881">Contact</span><span class="sxs-lookup"><span data-stu-id="ab066-881">Contact</span></span> | <span data-ttu-id="ab066-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ab066-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ab066-883">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-883">String</span></span> | <span data-ttu-id="ab066-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ab066-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ab066-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ab066-885">MeetingSuggestion</span></span> | <span data-ttu-id="ab066-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ab066-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ab066-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ab066-887">PhoneNumber</span></span> | <span data-ttu-id="ab066-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ab066-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ab066-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ab066-889">TaskSuggestion</span></span> | <span data-ttu-id="ab066-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ab066-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ab066-891">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-891">String</span></span> | <span data-ttu-id="ab066-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ab066-892">**Restricted**</span></span> |

<span data-ttu-id="ab066-893">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ab066-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="ab066-894">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-894">Example</span></span>

<span data-ttu-id="ab066-895">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa os endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="ab066-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="ab066-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ab066-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ab066-897">Retorna entidades conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ab066-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-898">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-899">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor especificado no elemento `FilterName` .</span><span class="sxs-lookup"><span data-stu-id="ab066-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-900">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-900">Parameters:</span></span>

|<span data-ttu-id="ab066-901">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-901">Name</span></span>| <span data-ttu-id="ab066-902">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-902">Type</span></span>| <span data-ttu-id="ab066-903">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ab066-904">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-904">String</span></span>|<span data-ttu-id="ab066-905">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="ab066-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-906">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-906">Requirements</span></span>

|<span data-ttu-id="ab066-907">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-907">Requirement</span></span>| <span data-ttu-id="ab066-908">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-909">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-910">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-910">1.0</span></span>|
|[<span data-ttu-id="ab066-911">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-912">ReadItem</span></span>|
|[<span data-ttu-id="ab066-913">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-914">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-915">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-915">Returns:</span></span>

<span data-ttu-id="ab066-p155">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="ab066-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ab066-918">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ab066-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="ab066-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ab066-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ab066-920">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ab066-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-921">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-p156">O método `getRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="ab066-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ab066-925">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="ab066-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ab066-926">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ab066-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ab066-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="ab066-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-930">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-930">Requirements</span></span>

|<span data-ttu-id="ab066-931">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-931">Requirement</span></span>| <span data-ttu-id="ab066-932">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-933">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-934">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-934">1.0</span></span>|
|[<span data-ttu-id="ab066-935">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-936">ReadItem</span></span>|
|[<span data-ttu-id="ab066-937">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-938">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-939">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-939">Returns:</span></span>

<span data-ttu-id="ab066-p158">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="ab066-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ab066-942">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ab066-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ab066-943">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ab066-944">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-944">Example</span></span>

<span data-ttu-id="ab066-945">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="ab066-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ab066-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ab066-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ab066-947">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ab066-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-948">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-949">O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="ab066-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ab066-p159">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="ab066-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-952">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-952">Parameters:</span></span>

|<span data-ttu-id="ab066-953">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-953">Name</span></span>| <span data-ttu-id="ab066-954">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-954">Type</span></span>| <span data-ttu-id="ab066-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ab066-956">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-956">String</span></span>|<span data-ttu-id="ab066-957">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="ab066-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-958">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-958">Requirements</span></span>

|<span data-ttu-id="ab066-959">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-959">Requirement</span></span>| <span data-ttu-id="ab066-960">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-961">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-962">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-962">1.0</span></span>|
|[<span data-ttu-id="ab066-963">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-964">ReadItem</span></span>|
|[<span data-ttu-id="ab066-965">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-966">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-967">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-967">Returns:</span></span>

<span data-ttu-id="ab066-968">Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ab066-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ab066-969">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ab066-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ab066-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ab066-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ab066-971">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ab066-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ab066-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ab066-973">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ab066-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ab066-p160">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="ab066-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-976">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-976">Parameters:</span></span>

|<span data-ttu-id="ab066-977">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-977">Name</span></span>| <span data-ttu-id="ab066-978">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-978">Type</span></span>| <span data-ttu-id="ab066-979">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-979">Attributes</span></span>| <span data-ttu-id="ab066-980">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="ab066-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ab066-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ab066-p161">Solicita um formato para os dados. Se for Text, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="ab066-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="ab066-985">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-985">Object</span></span>| <span data-ttu-id="ab066-986">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-986">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-987">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ab066-988">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-988">Object</span></span>| <span data-ttu-id="ab066-989">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-989">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-990">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ab066-991">função</span><span class="sxs-lookup"><span data-stu-id="ab066-991">function</span></span>||<span data-ttu-id="ab066-992">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ab066-993">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="ab066-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ab066-994">Para acessar a propriedade de origem de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="ab066-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-995">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-995">Requirements</span></span>

|<span data-ttu-id="ab066-996">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-996">Requirement</span></span>| <span data-ttu-id="ab066-997">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-998">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-999">1.2</span><span class="sxs-lookup"><span data-stu-id="ab066-999">1.2</span></span>|
|[<span data-ttu-id="ab066-1000">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-1002">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1003">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-1004">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-1004">Returns:</span></span>

<span data-ttu-id="ab066-1005">Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="ab066-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="ab066-1006">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="ab066-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ab066-1007">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ab066-1008">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1008">Example</span></span>

```
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="ab066-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ab066-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="ab066-p163">Obtém as entidades encontradas em uma correspondência destacada que um usuário selecionou. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ab066-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-1012">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-1013">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1013">Requirements</span></span>

|<span data-ttu-id="ab066-1014">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1014">Requirement</span></span>| <span data-ttu-id="ab066-1015">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1016">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="ab066-1017">-16</span></span> |
|[<span data-ttu-id="ab066-1018">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1019">ReadItem</span></span>|
|[<span data-ttu-id="ab066-1020">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1021">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-1022">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-1022">Returns:</span></span>

<span data-ttu-id="ab066-1023">Tipo: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ab066-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ab066-1024">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1024">Example</span></span>

<span data-ttu-id="ab066-1025">O exemplo a seguir acessa as entidades de endereços na correspondência destacada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="ab066-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="ab066-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ab066-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="ab066-p164">Retorna valores do tipo sequência de caracteres em uma correspondência destacada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="ab066-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-1029">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="ab066-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ab066-p165">O método `getSelectedRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="ab066-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ab066-1033">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="ab066-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ab066-1034">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ab066-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ab066-p166">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="ab066-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab066-1038">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1038">Requirements</span></span>

|<span data-ttu-id="ab066-1039">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1039">Requirement</span></span>| <span data-ttu-id="ab066-1040">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1041">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="ab066-1042">-16</span></span> |
|[<span data-ttu-id="ab066-1043">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1044">ReadItem</span></span>|
|[<span data-ttu-id="ab066-1045">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1046">Leitura</span><span class="sxs-lookup"><span data-stu-id="ab066-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ab066-1047">Retorna:</span><span class="sxs-lookup"><span data-stu-id="ab066-1047">Returns:</span></span>

<span data-ttu-id="ab066-p167">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="ab066-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="ab066-1050">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1050">Example</span></span>

<span data-ttu-id="ab066-1051">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="ab066-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ab066-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ab066-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ab066-1053">Carrega de forma assíncrona as propriedades personalizadas desse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="ab066-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ab066-p168">As propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item e o suplemento atuais. As propriedades personalizadas não são criptografadas no item e, portanto, isso não deve ser usado como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="ab066-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-1057">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-1057">Parameters:</span></span>

|<span data-ttu-id="ab066-1058">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-1058">Name</span></span>| <span data-ttu-id="ab066-1059">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-1059">Type</span></span>| <span data-ttu-id="ab066-1060">Attributes</span><span class="sxs-lookup"><span data-stu-id="ab066-1060">Attributes</span></span>| <span data-ttu-id="ab066-1061">Description</span><span class="sxs-lookup"><span data-stu-id="ab066-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ab066-1062">função</span><span class="sxs-lookup"><span data-stu-id="ab066-1062">function</span></span>||<span data-ttu-id="ab066-1063">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ab066-1064">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ab066-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ab066-1065">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas no servidor.</span><span class="sxs-lookup"><span data-stu-id="ab066-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ab066-1066">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1066">Object</span></span>| <span data-ttu-id="ab066-1067">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1068">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="ab066-1069">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-1070">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1070">Requirements</span></span>

|<span data-ttu-id="ab066-1071">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1071">Requirement</span></span>| <span data-ttu-id="ab066-1072">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1073">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="ab066-1074">1.0</span></span>|
|[<span data-ttu-id="ab066-1075">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1076">ReadItem</span></span>|
|[<span data-ttu-id="ab066-1077">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1078">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="ab066-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-1079">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1079">Example</span></span>

<span data-ttu-id="ab066-p171">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas específicas do item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ab066-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ab066-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ab066-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ab066-1084">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="ab066-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ab066-p172">O método `removeAttachmentAsync` remove do item o anexo com o identificador especificado. Conforme as práticas recomendadas, você deve usar o identificador do anexo para remover o anexo apenas se o mesmo aplicativo de email tiver inserido o anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador de anexos é válido somente dentro da mesma sessão. Uma sessão é considerada encerrada quando o usuário fecha o aplicativo, ou se o usuário começa a escrever um email em um formulário embutido e, em seguida, abre o mesmo formulário em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="ab066-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-1089">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-1089">Parameters:</span></span>

|<span data-ttu-id="ab066-1090">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-1090">Name</span></span>| <span data-ttu-id="ab066-1091">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-1091">Type</span></span>| <span data-ttu-id="ab066-1092">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-1092">Attributes</span></span>| <span data-ttu-id="ab066-1093">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ab066-1094">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-1094">String</span></span>||<span data-ttu-id="ab066-p173">O identificador do anexo a ser removido. O comprimento máximo da sequência de caracteres é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ab066-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="ab066-1097">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1097">Object</span></span>| <span data-ttu-id="ab066-1098">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1099">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ab066-1100">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1100">Object</span></span>| <span data-ttu-id="ab066-1101">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1102">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ab066-1103">função</span><span class="sxs-lookup"><span data-stu-id="ab066-1103">function</span></span>| <span data-ttu-id="ab066-1104">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1105">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ab066-1106">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="ab066-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ab066-1107">Erros</span><span class="sxs-lookup"><span data-stu-id="ab066-1107">Errors</span></span>

| <span data-ttu-id="ab066-1108">Código de erro</span><span class="sxs-lookup"><span data-stu-id="ab066-1108">Error code</span></span> | <span data-ttu-id="ab066-1109">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ab066-1110">O identificador do anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="ab066-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-1111">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1111">Requirements</span></span>

|<span data-ttu-id="ab066-1112">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1112">Requirement</span></span>| <span data-ttu-id="ab066-1113">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1114">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="ab066-1115">1.1</span></span>|
|[<span data-ttu-id="ab066-1116">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-1118">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1119">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-1120">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1120">Example</span></span>

<span data-ttu-id="ab066-1121">O código a seguir remove um anexo com um identificador igual a '0'.</span><span class="sxs-lookup"><span data-stu-id="ab066-1121">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="ab066-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ab066-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="ab066-1123">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ab066-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="ab066-p174">Quando chamado, este método salva a mensagem atual como um rascunho e retorna o identificador do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook em modo de cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="ab066-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-1127">Se o seu suplemento chamar `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver em modo de cache, poderá levar algum tempo antes do item realmente ser sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="ab066-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="ab066-1128">Até que o item seja sincronizado, o uso de `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="ab066-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ab066-p176">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo de redação, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="ab066-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ab066-1132">Os seguintes clientes possuem um comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="ab066-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ab066-1133">O Outlook para Mac não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ab066-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="ab066-1134">Chamar `saveAsync` em uma reunião no Outlook para Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="ab066-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="ab066-1135">O Outlook na Web sempre enviará um convite ou atualização quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="ab066-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-1136">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-1136">Parameters:</span></span>

|<span data-ttu-id="ab066-1137">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-1137">Name</span></span>| <span data-ttu-id="ab066-1138">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-1138">Type</span></span>| <span data-ttu-id="ab066-1139">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-1139">Attributes</span></span>| <span data-ttu-id="ab066-1140">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="ab066-1141">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1141">Object</span></span>| <span data-ttu-id="ab066-1142">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1143">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ab066-1144">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1144">Object</span></span>| <span data-ttu-id="ab066-1145">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1146">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ab066-1147">função</span><span class="sxs-lookup"><span data-stu-id="ab066-1147">function</span></span>||<span data-ttu-id="ab066-1148">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ab066-1149">Em caso de sucesso, o identificador do item será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ab066-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab066-1150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1150">Requirements</span></span>

|<span data-ttu-id="ab066-1151">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1151">Requirement</span></span>| <span data-ttu-id="ab066-1152">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1153">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="ab066-1154">1.3</span></span>|
|[<span data-ttu-id="ab066-1155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-1157">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1158">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ab066-1159">Exemplos</span><span class="sxs-lookup"><span data-stu-id="ab066-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="ab066-p178">A seguir apresentamos um exemplo do parâmetro `result` passado para a função de retorno de chamada. A propriedade `value` contém o ID do item.</span><span class="sxs-lookup"><span data-stu-id="ab066-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ab066-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ab066-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ab066-1163">Insere dados no corpo ou no assunto de uma mensagem de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ab066-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ab066-p179">O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou no assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou do assunto, um erro será retornado. Após a inserção, o cursor será posicionado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="ab066-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ab066-1167">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="ab066-1167">Parameters:</span></span>

|<span data-ttu-id="ab066-1168">Name</span><span class="sxs-lookup"><span data-stu-id="ab066-1168">Name</span></span>| <span data-ttu-id="ab066-1169">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab066-1169">Type</span></span>| <span data-ttu-id="ab066-1170">Atributos</span><span class="sxs-lookup"><span data-stu-id="ab066-1170">Attributes</span></span>| <span data-ttu-id="ab066-1171">Descrição</span><span class="sxs-lookup"><span data-stu-id="ab066-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ab066-1172">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="ab066-1172">String</span></span>||<span data-ttu-id="ab066-p180">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="ab066-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="ab066-1176">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1176">Object</span></span>| <span data-ttu-id="ab066-1177">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1178">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="ab066-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ab066-1179">Objeto</span><span class="sxs-lookup"><span data-stu-id="ab066-1179">Object</span></span>| <span data-ttu-id="ab066-1180">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-1181">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="ab066-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="ab066-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ab066-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="ab066-1183">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="ab066-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="ab066-p181">Se for `text` , o estilo atual será aplicado no Outlook Web App e no Outlook. Se o campo for um editor HTML, somente os dados de texto serão inseridos, mesmo que os dados estejam em HTML.</span><span class="sxs-lookup"><span data-stu-id="ab066-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ab066-p182">Se for `html` e o campo for compatível com HTML (e o assunto não), o estilo atual será aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, um erro `InvalidDataFormat` será retornado.</span><span class="sxs-lookup"><span data-stu-id="ab066-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ab066-1188">Se `coercionType` não estiver definido, o resultado dependerá do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="ab066-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="ab066-1189">função</span><span class="sxs-lookup"><span data-stu-id="ab066-1189">function</span></span>||<span data-ttu-id="ab066-1190">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ab066-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ab066-1191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab066-1191">Requirements</span></span>

|<span data-ttu-id="ab066-1192">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab066-1192">Requirement</span></span>| <span data-ttu-id="ab066-1193">Valor</span><span class="sxs-lookup"><span data-stu-id="ab066-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab066-1194">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab066-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab066-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="ab066-1195">1.2</span></span>|
|[<span data-ttu-id="ab066-1196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab066-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab066-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ab066-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="ab066-1198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab066-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ab066-1199">Redigir</span><span class="sxs-lookup"><span data-stu-id="ab066-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ab066-1200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab066-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```