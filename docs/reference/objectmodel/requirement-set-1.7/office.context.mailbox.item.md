
# <a name="item"></a><span data-ttu-id="1f699-101">item</span><span class="sxs-lookup"><span data-stu-id="1f699-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1f699-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1f699-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1f699-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1f699-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-105">Requirements</span></span>

|<span data-ttu-id="1f699-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-106">Requirement</span></span>|<span data-ttu-id="1f699-107">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-109">1.0</span></span>|
|[<span data-ttu-id="1f699-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="1f699-111">Restricted</span></span>|
|[<span data-ttu-id="1f699-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1f699-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="1f699-114">Members and methods</span></span>

| <span data-ttu-id="1f699-115">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-115">Member</span></span> | <span data-ttu-id="1f699-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1f699-117">attachments</span><span class="sxs-lookup"><span data-stu-id="1f699-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="1f699-118">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-118">Member</span></span> |
| [<span data-ttu-id="1f699-119">bcc</span><span class="sxs-lookup"><span data-stu-id="1f699-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1f699-120">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-120">Member</span></span> |
| [<span data-ttu-id="1f699-121">body</span><span class="sxs-lookup"><span data-stu-id="1f699-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="1f699-122">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-122">Member</span></span> |
| [<span data-ttu-id="1f699-123">cc</span><span class="sxs-lookup"><span data-stu-id="1f699-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1f699-124">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-124">Member</span></span> |
| [<span data-ttu-id="1f699-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="1f699-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1f699-126">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-126">Member</span></span> |
| [<span data-ttu-id="1f699-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1f699-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1f699-128">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-128">Member</span></span> |
| [<span data-ttu-id="1f699-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1f699-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1f699-130">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-130">Member</span></span> |
| [<span data-ttu-id="1f699-131">end</span><span class="sxs-lookup"><span data-stu-id="1f699-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="1f699-132">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-132">Member</span></span> |
| [<span data-ttu-id="1f699-133">from</span><span class="sxs-lookup"><span data-stu-id="1f699-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="1f699-134">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-134">Member</span></span> |
| [<span data-ttu-id="1f699-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1f699-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1f699-136">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-136">Member</span></span> |
| [<span data-ttu-id="1f699-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="1f699-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1f699-138">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-138">Member</span></span> |
| [<span data-ttu-id="1f699-139">itemId</span><span class="sxs-lookup"><span data-stu-id="1f699-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1f699-140">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-140">Member</span></span> |
| [<span data-ttu-id="1f699-141">itemType</span><span class="sxs-lookup"><span data-stu-id="1f699-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="1f699-142">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-142">Member</span></span> |
| [<span data-ttu-id="1f699-143">location</span><span class="sxs-lookup"><span data-stu-id="1f699-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="1f699-144">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-144">Member</span></span> |
| [<span data-ttu-id="1f699-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1f699-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1f699-146">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-146">Member</span></span> |
| [<span data-ttu-id="1f699-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1f699-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="1f699-148">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-148">Member</span></span> |
| [<span data-ttu-id="1f699-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1f699-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1f699-150">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-150">Member</span></span> |
| [<span data-ttu-id="1f699-151">organizer</span><span class="sxs-lookup"><span data-stu-id="1f699-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="1f699-152">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-152">Member</span></span> |
| [<span data-ttu-id="1f699-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="1f699-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="1f699-154">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-154">Member</span></span> |
| [<span data-ttu-id="1f699-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1f699-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1f699-156">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-156">Member</span></span> |
| [<span data-ttu-id="1f699-157">sender</span><span class="sxs-lookup"><span data-stu-id="1f699-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="1f699-158">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-158">Member</span></span> |
| [<span data-ttu-id="1f699-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="1f699-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="1f699-160">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-160">Member</span></span> |
| [<span data-ttu-id="1f699-161">start</span><span class="sxs-lookup"><span data-stu-id="1f699-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="1f699-162">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-162">Member</span></span> |
| [<span data-ttu-id="1f699-163">subject</span><span class="sxs-lookup"><span data-stu-id="1f699-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="1f699-164">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-164">Member</span></span> |
| [<span data-ttu-id="1f699-165">to</span><span class="sxs-lookup"><span data-stu-id="1f699-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1f699-166">Membro</span><span class="sxs-lookup"><span data-stu-id="1f699-166">Member</span></span> |
| [<span data-ttu-id="1f699-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1f699-168">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-168">Method</span></span> |
| [<span data-ttu-id="1f699-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1f699-170">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-170">Method</span></span> |
| [<span data-ttu-id="1f699-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1f699-172">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-172">Method</span></span> |
| [<span data-ttu-id="1f699-173">close</span><span class="sxs-lookup"><span data-stu-id="1f699-173">close</span></span>](#close) | <span data-ttu-id="1f699-174">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-174">Method</span></span> |
| [<span data-ttu-id="1f699-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1f699-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="1f699-176">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-176">Method</span></span> |
| [<span data-ttu-id="1f699-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1f699-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="1f699-178">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-178">Method</span></span> |
| [<span data-ttu-id="1f699-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="1f699-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="1f699-180">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-180">Method</span></span> |
| [<span data-ttu-id="1f699-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1f699-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="1f699-182">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-182">Method</span></span> |
| [<span data-ttu-id="1f699-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1f699-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="1f699-184">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-184">Method</span></span> |
| [<span data-ttu-id="1f699-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1f699-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1f699-186">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-186">Method</span></span> |
| [<span data-ttu-id="1f699-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1f699-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1f699-188">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-188">Method</span></span> |
| [<span data-ttu-id="1f699-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1f699-190">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-190">Method</span></span> |
| [<span data-ttu-id="1f699-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="1f699-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="1f699-192">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-192">Method</span></span> |
| [<span data-ttu-id="1f699-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1f699-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1f699-194">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-194">Method</span></span> |
| [<span data-ttu-id="1f699-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1f699-196">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-196">Method</span></span> |
| [<span data-ttu-id="1f699-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1f699-198">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-198">Method</span></span> |
| [<span data-ttu-id="1f699-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1f699-200">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-200">Method</span></span> |
| [<span data-ttu-id="1f699-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1f699-202">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-202">Method</span></span> |
| [<span data-ttu-id="1f699-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1f699-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1f699-204">Método</span><span class="sxs-lookup"><span data-stu-id="1f699-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1f699-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-205">Example</span></span>

<span data-ttu-id="1f699-206">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="1f699-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1f699-207">Membros</span><span class="sxs-lookup"><span data-stu-id="1f699-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="1f699-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1f699-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="1f699-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-211">Certos tipos de arquivos são bloqueados pelo Outlook devido a potenciais problemas de segurança e portanto não são retornados.</span><span class="sxs-lookup"><span data-stu-id="1f699-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1f699-212">Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1f699-212">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-213">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-213">Type:</span></span>

*   <span data-ttu-id="1f699-214">Array. <[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1f699-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-215">Requirements</span></span>

|<span data-ttu-id="1f699-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-216">Requirement</span></span>|<span data-ttu-id="1f699-217">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-218">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-219">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-219">1.0</span></span>|
|[<span data-ttu-id="1f699-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-221">ReadItem</span></span>|
|[<span data-ttu-id="1f699-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-223">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-224">Example</span></span>

<span data-ttu-id="1f699-225">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1f699-226">cco:[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1f699-227">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1f699-228">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="1f699-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-229">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-229">Type:</span></span>

*   [<span data-ttu-id="1f699-230">Destinatários</span><span class="sxs-lookup"><span data-stu-id="1f699-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1f699-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-231">Requirements</span></span>

|<span data-ttu-id="1f699-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-232">Requirement</span></span>|<span data-ttu-id="1f699-233">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-234">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-234">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-235">1.1</span><span class="sxs-lookup"><span data-stu-id="1f699-235">1.1</span></span>|
|[<span data-ttu-id="1f699-236">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-237">ReadItem</span></span>|
|[<span data-ttu-id="1f699-238">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-239">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-240">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="1f699-241">corpo:[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="1f699-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="1f699-242">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="1f699-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-243">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-243">Type:</span></span>

*   [<span data-ttu-id="1f699-244">Body</span><span class="sxs-lookup"><span data-stu-id="1f699-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="1f699-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-245">Requirements</span></span>

|<span data-ttu-id="1f699-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-246">Requirement</span></span>|<span data-ttu-id="1f699-247">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-248">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-249">1.1</span><span class="sxs-lookup"><span data-stu-id="1f699-249">1.1</span></span>|
|[<span data-ttu-id="1f699-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-251">ReadItem</span></span>|
|[<span data-ttu-id="1f699-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-253">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1f699-254">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1f699-255">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1f699-256">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-257">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-257">Read mode</span></span>

<span data-ttu-id="1f699-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1f699-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-260">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-260">Compose mode</span></span>

<span data-ttu-id="1f699-261">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-261">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-262">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-262">Type:</span></span>

*   <span data-ttu-id="1f699-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-264">Requirements</span></span>

|<span data-ttu-id="1f699-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-265">Requirement</span></span>|<span data-ttu-id="1f699-266">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-267">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-268">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-268">1.0</span></span>|
|[<span data-ttu-id="1f699-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-270">ReadItem</span></span>|
|[<span data-ttu-id="1f699-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-272">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-273">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1f699-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1f699-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="1f699-275">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="1f699-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1f699-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviá-la, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não será mais aplicável.</span><span class="sxs-lookup"><span data-stu-id="1f699-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1f699-p108">Para um novo item em um formulário de redação, o valor dessa propriedade é nulo. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="1f699-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-280">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-280">Type:</span></span>

*   <span data-ttu-id="1f699-281">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-282">Requirements</span></span>

|<span data-ttu-id="1f699-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-283">Requirement</span></span>|<span data-ttu-id="1f699-284">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-285">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-286">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-286">1.0</span></span>|
|[<span data-ttu-id="1f699-287">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-288">ReadItem</span></span>|
|[<span data-ttu-id="1f699-289">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-290">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1f699-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1f699-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="1f699-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-294">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-294">Type:</span></span>

*   <span data-ttu-id="1f699-295">Data</span><span class="sxs-lookup"><span data-stu-id="1f699-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-296">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-296">Requirements</span></span>

|<span data-ttu-id="1f699-297">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-297">Requirement</span></span>|<span data-ttu-id="1f699-298">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-299">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-300">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-300">1.0</span></span>|
|[<span data-ttu-id="1f699-301">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-302">ReadItem</span></span>|
|[<span data-ttu-id="1f699-303">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-304">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-305">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1f699-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1f699-306">dateTimeModified :Date</span></span>

<span data-ttu-id="1f699-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-309">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-309">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-310">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-310">Type:</span></span>

*   <span data-ttu-id="1f699-311">Data</span><span class="sxs-lookup"><span data-stu-id="1f699-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-312">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-312">Requirements</span></span>

|<span data-ttu-id="1f699-313">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-313">Requirement</span></span>|<span data-ttu-id="1f699-314">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-315">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-316">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-316">1.0</span></span>|
|[<span data-ttu-id="1f699-317">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-318">ReadItem</span></span>|
|[<span data-ttu-id="1f699-319">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-320">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-321">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1f699-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1f699-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1f699-323">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="1f699-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1f699-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor da propriedade para a data e hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1f699-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-326">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-326">Read mode</span></span>

<span data-ttu-id="1f699-327">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1f699-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-328">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-328">Compose mode</span></span>

<span data-ttu-id="1f699-329">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1f699-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1f699-330">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC do servidor.</span><span class="sxs-lookup"><span data-stu-id="1f699-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-331">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-331">Type:</span></span>

*   <span data-ttu-id="1f699-332">Data | [Hora](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1f699-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-333">Requirements</span></span>

|<span data-ttu-id="1f699-334">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-334">Requirement</span></span>|<span data-ttu-id="1f699-335">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-336">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-337">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-337">1.0</span></span>|
|[<span data-ttu-id="1f699-338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-339">ReadItem</span></span>|
|[<span data-ttu-id="1f699-340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-341">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-342">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-342">Example</span></span>

<span data-ttu-id="1f699-343">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1f699-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="1f699-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="1f699-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="1f699-345">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="1f699-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="1f699-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-348">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1f699-348">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-349">Read mode</span></span>

<span data-ttu-id="1f699-350">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="1f699-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="1f699-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-351">Compose mode</span></span>

<span data-ttu-id="1f699-352">A propriedade `from` retornará um objeto `From` que fornece um método para obter o valor de from.</span><span class="sxs-lookup"><span data-stu-id="1f699-352">Added From: Adds a new object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1f699-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-353">Type:</span></span>

*   <span data-ttu-id="1f699-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="1f699-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-355">Requirements</span></span>

|<span data-ttu-id="1f699-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1f699-357">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-358">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-358">1.0</span></span>|<span data-ttu-id="1f699-359">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-359">-17</span></span>|
|[<span data-ttu-id="1f699-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-361">ReadItem</span></span>|<span data-ttu-id="1f699-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-363">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-364">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-364">Read</span></span>|<span data-ttu-id="1f699-365">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1f699-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1f699-366">internetMessageId :String</span></span>

<span data-ttu-id="1f699-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-369">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-369">Type:</span></span>

*   <span data-ttu-id="1f699-370">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-371">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-371">Requirements</span></span>

|<span data-ttu-id="1f699-372">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-372">Requirement</span></span>|<span data-ttu-id="1f699-373">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-374">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-375">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-375">1.0</span></span>|
|[<span data-ttu-id="1f699-376">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-377">ReadItem</span></span>|
|[<span data-ttu-id="1f699-378">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-379">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-380">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1f699-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1f699-381">itemClass :String</span></span>

<span data-ttu-id="1f699-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1f699-p115">A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir estão as classes de mensagem padrão para itens de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="1f699-386">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-386">Type</span></span>|<span data-ttu-id="1f699-387">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-387">Description</span></span>|<span data-ttu-id="1f699-388">classe do item</span><span class="sxs-lookup"><span data-stu-id="1f699-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="1f699-389">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="1f699-389">Appointment items</span></span>|<span data-ttu-id="1f699-390">São itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="1f699-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="1f699-391">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="1f699-391">Message items</span></span>|<span data-ttu-id="1f699-392">Incluem mensagens de e-mail que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagens base.</span><span class="sxs-lookup"><span data-stu-id="1f699-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="1f699-393">Você pode criar classes de mensagens personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="1f699-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-394">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-394">Type:</span></span>

*   <span data-ttu-id="1f699-395">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-396">Requirements</span></span>

|<span data-ttu-id="1f699-397">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-397">Requirement</span></span>|<span data-ttu-id="1f699-398">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-399">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-400">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-400">1.0</span></span>|
|[<span data-ttu-id="1f699-401">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-402">ReadItem</span></span>|
|[<span data-ttu-id="1f699-403">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-404">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-405">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1f699-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1f699-406">(nullable) itemId :String</span></span>

<span data-ttu-id="1f699-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-409">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1f699-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1f699-410">A propriedade `itemId` não é idêntica ao Entry ID do Outlook ou ao ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1f699-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1f699-411">Antes de fazer chamadas à API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="1f699-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1f699-412">Para mais informações, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1f699-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1f699-p118">A propriedade `itemId` não está disponível no modo de redação. Se  um identificador de item for requerido, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no repositório, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-415">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-415">Type:</span></span>

*   <span data-ttu-id="1f699-416">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-417">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-417">Requirements</span></span>

|<span data-ttu-id="1f699-418">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-418">Requirement</span></span>|<span data-ttu-id="1f699-419">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-420">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-421">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-421">1.0</span></span>|
|[<span data-ttu-id="1f699-422">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-423">ReadItem</span></span>|
|[<span data-ttu-id="1f699-424">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-425">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-426">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-426">Example</span></span>

<span data-ttu-id="1f699-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item a partir do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1f699-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="1f699-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1f699-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1f699-430">Obtém o tipo de item que uma instância representa.</span><span class="sxs-lookup"><span data-stu-id="1f699-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1f699-431">A propriedade `itemType` retorna um dos valores da enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-432">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-432">Type:</span></span>

*   [<span data-ttu-id="1f699-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1f699-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1f699-434">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-434">Requirements</span></span>

|<span data-ttu-id="1f699-435">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-435">Requirement</span></span>|<span data-ttu-id="1f699-436">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-437">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-438">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-438">1.0</span></span>|
|[<span data-ttu-id="1f699-439">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-440">ReadItem</span></span>|
|[<span data-ttu-id="1f699-441">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-442">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-443">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="1f699-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1f699-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="1f699-445">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-446">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-446">Read mode</span></span>

<span data-ttu-id="1f699-447">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-448">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-448">Compose mode</span></span>

<span data-ttu-id="1f699-449">A propriedade `location` retorna um objeto `Location` que fornece métodos para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-450">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-450">Type:</span></span>

*   <span data-ttu-id="1f699-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1f699-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-452">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-452">Requirements</span></span>

|<span data-ttu-id="1f699-453">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-453">Requirement</span></span>|<span data-ttu-id="1f699-454">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-455">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-456">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-456">1.0</span></span>|
|[<span data-ttu-id="1f699-457">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-458">ReadItem</span></span>|
|[<span data-ttu-id="1f699-459">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-460">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-461">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1f699-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1f699-462">normalizedSubject :String</span></span>

<span data-ttu-id="1f699-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1f699-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="1f699-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-467">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-467">Type:</span></span>

*   <span data-ttu-id="1f699-468">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-469">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-469">Requirements</span></span>

|<span data-ttu-id="1f699-470">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-470">Requirement</span></span>|<span data-ttu-id="1f699-471">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-472">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-473">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-473">1.0</span></span>|
|[<span data-ttu-id="1f699-474">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-475">ReadItem</span></span>|
|[<span data-ttu-id="1f699-476">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-477">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-478">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="1f699-479">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1f699-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="1f699-480">Obtém as mensagens de notificação para um item.</span><span class="sxs-lookup"><span data-stu-id="1f699-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-481">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-481">Type:</span></span>

*   [<span data-ttu-id="1f699-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1f699-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1f699-483">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-483">Requirements</span></span>

|<span data-ttu-id="1f699-484">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-484">Requirement</span></span>|<span data-ttu-id="1f699-485">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-486">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-487">1.3</span><span class="sxs-lookup"><span data-stu-id="1f699-487">1.3</span></span>|
|[<span data-ttu-id="1f699-488">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-489">ReadItem</span></span>|
|[<span data-ttu-id="1f699-490">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-491">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1f699-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1f699-493">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="1f699-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1f699-494">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-495">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-495">Read mode</span></span>

<span data-ttu-id="1f699-496">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-497">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-497">Compose mode</span></span>

<span data-ttu-id="1f699-498">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-499">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-499">Type:</span></span>

*   <span data-ttu-id="1f699-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-501">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-501">Requirements</span></span>

|<span data-ttu-id="1f699-502">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-502">Requirement</span></span>|<span data-ttu-id="1f699-503">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-504">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-505">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-505">1.0</span></span>|
|[<span data-ttu-id="1f699-506">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-507">ReadItem</span></span>|
|[<span data-ttu-id="1f699-508">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-509">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-510">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="1f699-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1f699-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="1f699-512">Obtém o endereço de email do organizador da reunião para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="1f699-512">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-513">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-513">Read mode</span></span>

<span data-ttu-id="1f699-514">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-515">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-515">Compose mode</span></span>

<span data-ttu-id="1f699-516">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook_1_7/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="1f699-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-517">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-517">Type:</span></span>

*   <span data-ttu-id="1f699-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1f699-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-519">Requirements</span></span>

|<span data-ttu-id="1f699-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1f699-521">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-522">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-522">1.0</span></span>|<span data-ttu-id="1f699-523">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-523">-17</span></span>|
|[<span data-ttu-id="1f699-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-525">ReadItem</span></span>|<span data-ttu-id="1f699-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-527">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-528">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-528">Read</span></span>|<span data-ttu-id="1f699-529">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-530">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="1f699-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="1f699-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="1f699-532">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-532">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="1f699-533">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="1f699-534">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="1f699-535">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="1f699-536">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook_1_7/office.recurrence) para solicitações de reuniões ou compromissos recorrentes, se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="1f699-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="1f699-537">`null` é retornada para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="1f699-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="1f699-538">`undefined` é retornado para mensagens que não fazem solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="1f699-539">Observação: Solicitações de reunião tem um valor `itemClass` de IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="1f699-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="1f699-540">Observação: Se o objeto de recorrência for `null`, isto indica que o objeto é um compromisso único ou uma solicitação de reunião de um compromisso único e NÃO faz parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="1f699-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-541">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-541">Type:</span></span>

* [<span data-ttu-id="1f699-542">Recorrência</span><span class="sxs-lookup"><span data-stu-id="1f699-542">recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="1f699-543">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-543">Requirement</span></span>|<span data-ttu-id="1f699-544">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-545">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-546">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-546">-17</span></span>|
|[<span data-ttu-id="1f699-547">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-548">ReadItem</span></span>|
|[<span data-ttu-id="1f699-549">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-550">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1f699-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1f699-552">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="1f699-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1f699-553">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-554">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-554">Read mode</span></span>

<span data-ttu-id="1f699-555">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-556">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-556">Compose mode</span></span>

<span data-ttu-id="1f699-557">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-558">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-558">Type:</span></span>

*   <span data-ttu-id="1f699-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-560">Requirements</span></span>

|<span data-ttu-id="1f699-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-561">Requirement</span></span>|<span data-ttu-id="1f699-562">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-563">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-564">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-564">1.0</span></span>|
|[<span data-ttu-id="1f699-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-566">ReadItem</span></span>|
|[<span data-ttu-id="1f699-567">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-568">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-569">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="1f699-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1f699-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="1f699-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1f699-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1f699-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegador, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="1f699-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-575">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1f699-575">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-576">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-576">Type:</span></span>

*   [<span data-ttu-id="1f699-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1f699-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1f699-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-578">Requirements</span></span>

|<span data-ttu-id="1f699-579">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-579">Requirement</span></span>|<span data-ttu-id="1f699-580">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-581">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-582">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-582">1.0</span></span>|
|[<span data-ttu-id="1f699-583">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-584">ReadItem</span></span>|
|[<span data-ttu-id="1f699-585">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-586">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-587">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="1f699-588">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="1f699-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="1f699-589">Obtém a identificação da série a que uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="1f699-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="1f699-590">No OWA e no Outlook, o `seriesId` retornará a identificação de serviços Web do Exchange (EWS) do item pai (série) a que este item pertence.</span><span class="sxs-lookup"><span data-stu-id="1f699-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="1f699-591">No entanto, em iOS e Android, o `seriesId` retornará a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="1f699-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-592">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1f699-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1f699-593">A propriedade `seriesId` não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1f699-593">The `seriesId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1f699-594">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="1f699-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1f699-595">Para mais informações, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="1f699-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="1f699-596">A  `seriesId`propriedade retornará `null`para itens que não têm itens pai como compromissos, itens de série, ou solicitações de reunião únicos e retorna `undefined` para todos os itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="1f699-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-597">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-597">Type:</span></span>

* <span data-ttu-id="1f699-598">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-599">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-599">Requirements</span></span>

|<span data-ttu-id="1f699-600">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-600">Requirement</span></span>|<span data-ttu-id="1f699-601">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-602">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-603">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-603">-17</span></span>|
|[<span data-ttu-id="1f699-604">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-605">ReadItem</span></span>|
|[<span data-ttu-id="1f699-606">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-607">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-608">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1f699-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1f699-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1f699-610">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="1f699-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1f699-p130">A propriedade `start` é expressa como um valor de data e valor temporal no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1f699-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-613">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-613">Read mode</span></span>

<span data-ttu-id="1f699-614">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="1f699-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-615">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-615">Compose mode</span></span>

<span data-ttu-id="1f699-616">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1f699-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1f699-617">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="1f699-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-618">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-618">Type:</span></span>

*   <span data-ttu-id="1f699-619">Data | [Hora](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1f699-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-620">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-620">Requirements</span></span>

|<span data-ttu-id="1f699-621">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-621">Requirement</span></span>|<span data-ttu-id="1f699-622">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-623">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-624">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-624">1.0</span></span>|
|[<span data-ttu-id="1f699-625">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-626">ReadItem</span></span>|
|[<span data-ttu-id="1f699-627">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-628">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-629">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-629">Example</span></span>

<span data-ttu-id="1f699-630">O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="1f699-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="1f699-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1f699-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="1f699-632">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="1f699-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1f699-633">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de e-mail.</span><span class="sxs-lookup"><span data-stu-id="1f699-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-634">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-634">Read mode</span></span>

<span data-ttu-id="1f699-p131">A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1f699-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1f699-637">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-637">Compose mode</span></span>

<span data-ttu-id="1f699-638">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="1f699-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1f699-639">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-639">Type:</span></span>

*   <span data-ttu-id="1f699-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1f699-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-641">Requirements</span></span>

|<span data-ttu-id="1f699-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-642">Requirement</span></span>|<span data-ttu-id="1f699-643">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-644">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-645">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-645">1.0</span></span>|
|[<span data-ttu-id="1f699-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-647">ReadItem</span></span>|
|[<span data-ttu-id="1f699-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-649">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1f699-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1f699-651">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1f699-652">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1f699-653">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-653">Read mode</span></span>

<span data-ttu-id="1f699-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **To** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="1f699-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1f699-656">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="1f699-656">Compose mode</span></span>

<span data-ttu-id="1f699-657">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **To** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-657">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1f699-658">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f699-658">Type:</span></span>

*   <span data-ttu-id="1f699-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1f699-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-660">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-660">Requirements</span></span>

|<span data-ttu-id="1f699-661">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-661">Requirement</span></span>|<span data-ttu-id="1f699-662">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-663">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-663">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-664">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-664">1.0</span></span>|
|[<span data-ttu-id="1f699-665">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-666">ReadItem</span></span>|
|[<span data-ttu-id="1f699-667">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-668">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-669">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1f699-670">Métodos</span><span class="sxs-lookup"><span data-stu-id="1f699-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1f699-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1f699-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1f699-672">Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.</span><span class="sxs-lookup"><span data-stu-id="1f699-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1f699-673">O método `addFileAttachmentAsync` carrega o arquivo da URI especificada e o anexa ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="1f699-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1f699-674">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1f699-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-675">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-675">Parameters:</span></span>
|<span data-ttu-id="1f699-676">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-676">Name</span></span>|<span data-ttu-id="1f699-677">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-677">Type</span></span>|<span data-ttu-id="1f699-678">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-678">Attributes</span></span>|<span data-ttu-id="1f699-679">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="1f699-680">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-680">String</span></span>||<span data-ttu-id="1f699-p134">O URI que fornece a localização do arquivo anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1f699-683">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-683">String</span></span>||<span data-ttu-id="1f699-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1f699-686">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-686">Object</span></span>|<span data-ttu-id="1f699-687">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-687">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-688">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-689">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-689">Object</span></span>|<span data-ttu-id="1f699-690">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-690">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-691">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="1f699-692">Booleano</span><span class="sxs-lookup"><span data-stu-id="1f699-692">Boolean</span></span>|<span data-ttu-id="1f699-693">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-693">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-694">Se for `true`, indicará que o anexo será embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="1f699-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="1f699-695">function</span><span class="sxs-lookup"><span data-stu-id="1f699-695">function</span></span>|<span data-ttu-id="1f699-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-696">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-697">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1f699-698">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1f699-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1f699-699">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1f699-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1f699-700">Erros</span><span class="sxs-lookup"><span data-stu-id="1f699-700">Errors</span></span>

|<span data-ttu-id="1f699-701">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1f699-701">Error code</span></span>|<span data-ttu-id="1f699-702">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="1f699-703">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="1f699-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="1f699-704">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="1f699-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1f699-705">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1f699-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-706">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-706">Requirements</span></span>

|<span data-ttu-id="1f699-707">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-707">Requirement</span></span>|<span data-ttu-id="1f699-708">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-709">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-709">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-710">1.1</span><span class="sxs-lookup"><span data-stu-id="1f699-710">1.1</span></span>|
|[<span data-ttu-id="1f699-711">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-713">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-714">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1f699-715">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1f699-715">Examples</span></span>

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

<span data-ttu-id="1f699-716">O exemplo a seguir adiciona um arquivo de imagem como um anexo em linha e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1f699-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1f699-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1f699-718">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="1f699-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1f699-719">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged` , `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="1f699-719">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-720">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-720">Parameters:</span></span>

| <span data-ttu-id="1f699-721">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-721">Name</span></span> | <span data-ttu-id="1f699-722">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-722">Type</span></span> | <span data-ttu-id="1f699-723">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-723">Attributes</span></span> | <span data-ttu-id="1f699-724">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1f699-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1f699-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1f699-726">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="1f699-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1f699-727">Função</span><span class="sxs-lookup"><span data-stu-id="1f699-727">Function</span></span> || <span data-ttu-id="1f699-p136">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1f699-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1f699-731">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-731">Object</span></span> | <span data-ttu-id="1f699-732">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-732">&lt;optional&gt;</span></span> | <span data-ttu-id="1f699-733">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1f699-734">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-734">Object</span></span> | <span data-ttu-id="1f699-735">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-735">&lt;optional&gt;</span></span> | <span data-ttu-id="1f699-736">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1f699-737">function</span><span class="sxs-lookup"><span data-stu-id="1f699-737">function</span></span>| <span data-ttu-id="1f699-738">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-738">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-739">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-740">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-740">Requirements</span></span>

|<span data-ttu-id="1f699-741">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-741">Requirement</span></span>| <span data-ttu-id="1f699-742">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-743">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-743">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f699-744">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-744">-17</span></span> |
|[<span data-ttu-id="1f699-745">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f699-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-746">ReadItem</span></span> |
|[<span data-ttu-id="1f699-747">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f699-748">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="1f699-749">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1f699-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1f699-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1f699-751">Adiciona um item do Exchange, como uma mensagem, como um anexo à mensagem ou ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1f699-p137">O método `addItemAttachmentAsync` anexa o item com o identificador especificado do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="1f699-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1f699-755">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="1f699-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1f699-756">Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a outros itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="1f699-756">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-757">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-757">Parameters:</span></span>

|<span data-ttu-id="1f699-758">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-758">Name</span></span>|<span data-ttu-id="1f699-759">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-759">Type</span></span>|<span data-ttu-id="1f699-760">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-760">Attributes</span></span>|<span data-ttu-id="1f699-761">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="1f699-762">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-762">String</span></span>||<span data-ttu-id="1f699-p138">O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1f699-765">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-765">String</span></span>||<span data-ttu-id="1f699-p139">O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1f699-768">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-768">Object</span></span>|<span data-ttu-id="1f699-769">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-769">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-770">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-771">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-771">Object</span></span>|<span data-ttu-id="1f699-772">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-772">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-773">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1f699-774">function</span><span class="sxs-lookup"><span data-stu-id="1f699-774">function</span></span>|<span data-ttu-id="1f699-775">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-775">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-776">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1f699-777">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1f699-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1f699-778">Se não for possível adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` com a descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="1f699-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1f699-779">Erros</span><span class="sxs-lookup"><span data-stu-id="1f699-779">Errors</span></span>

|<span data-ttu-id="1f699-780">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1f699-780">Error code</span></span>|<span data-ttu-id="1f699-781">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1f699-782">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="1f699-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-783">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-783">Requirements</span></span>

|<span data-ttu-id="1f699-784">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-784">Requirement</span></span>|<span data-ttu-id="1f699-785">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-786">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-787">1.1</span><span class="sxs-lookup"><span data-stu-id="1f699-787">1.1</span></span>|
|[<span data-ttu-id="1f699-788">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-790">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-791">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-792">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-792">Example</span></span>

<span data-ttu-id="1f699-793">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1f699-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="1f699-794">close()</span><span class="sxs-lookup"><span data-stu-id="1f699-794">close()</span></span>

<span data-ttu-id="1f699-795">Fecha o item atual que está sendo redigido.</span><span class="sxs-lookup"><span data-stu-id="1f699-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1f699-p140">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item possuir alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação de fechamento.</span><span class="sxs-lookup"><span data-stu-id="1f699-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-798">No Outlook na Web, se o item for um compromisso e tiver sido salvo anteriormente usando `saveAsync`, será solicitado ao usuário para salvar, descartar ou cancelar, mesmo que nenhuma alteração tenha ocorrido após o item ter sido salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="1f699-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1f699-799">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="1f699-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-800">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-800">Requirements</span></span>

|<span data-ttu-id="1f699-801">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-801">Requirement</span></span>|<span data-ttu-id="1f699-802">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-803">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-803">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-804">1.3</span><span class="sxs-lookup"><span data-stu-id="1f699-804">1.3</span></span>|
|[<span data-ttu-id="1f699-805">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-806">Restrito</span><span class="sxs-lookup"><span data-stu-id="1f699-806">Restricted</span></span>|
|[<span data-ttu-id="1f699-807">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-808">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1f699-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1f699-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1f699-810">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1f699-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-811">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-811">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-812">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1f699-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1f699-813">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1f699-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1f699-p141">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="1f699-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-817">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-817">Parameters:</span></span>

|<span data-ttu-id="1f699-818">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-818">Name</span></span>|<span data-ttu-id="1f699-819">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-819">Type</span></span>|<span data-ttu-id="1f699-820">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-820">Attributes</span></span>|<span data-ttu-id="1f699-821">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1f699-822">String | Object</span><span class="sxs-lookup"><span data-stu-id="1f699-822">String &#124; Object</span></span>||<span data-ttu-id="1f699-p142">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1f699-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1f699-825">**OU**</span><span class="sxs-lookup"><span data-stu-id="1f699-825">**OR**</span></span><br/><span data-ttu-id="1f699-p143">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="1f699-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1f699-828">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-828">String</span></span>|<span data-ttu-id="1f699-829">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-829">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-p144">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1f699-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1f699-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1f699-833">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-833">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-834">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1f699-835">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-835">String</span></span>||<span data-ttu-id="1f699-p145">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1f699-838">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-838">String</span></span>||<span data-ttu-id="1f699-839">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="1f699-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1f699-840">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-840">String</span></span>||<span data-ttu-id="1f699-p146">Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1f699-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1f699-843">Booleano</span><span class="sxs-lookup"><span data-stu-id="1f699-843">Boolean</span></span>||<span data-ttu-id="1f699-p147">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="1f699-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1f699-846">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-846">String</span></span>||<span data-ttu-id="1f699-p148">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1f699-850">função</span><span class="sxs-lookup"><span data-stu-id="1f699-850">function</span></span>|<span data-ttu-id="1f699-851">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-851">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-852">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-853">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-853">Requirements</span></span>

|<span data-ttu-id="1f699-854">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-854">Requirement</span></span>|<span data-ttu-id="1f699-855">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-856">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-857">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-857">1.0</span></span>|
|[<span data-ttu-id="1f699-858">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-859">ReadItem</span></span>|
|[<span data-ttu-id="1f699-860">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-861">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1f699-862">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1f699-862">Examples</span></span>

<span data-ttu-id="1f699-863">O código a seguir passa uma sequência de caracteres para a função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1f699-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1f699-864">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1f699-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1f699-865">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1f699-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1f699-866">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="1f699-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1f699-867">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1f699-868">Resposta com o corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1f699-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1f699-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="1f699-870">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="1f699-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-871">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-871">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-872">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="1f699-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1f699-873">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1f699-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1f699-p149">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="1f699-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-877">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-877">Parameters:</span></span>

|<span data-ttu-id="1f699-878">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-878">Name</span></span>|<span data-ttu-id="1f699-879">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-879">Type</span></span>|<span data-ttu-id="1f699-880">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-880">Attributes</span></span>|<span data-ttu-id="1f699-881">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1f699-882">String | Object</span><span class="sxs-lookup"><span data-stu-id="1f699-882">String &#124; Object</span></span>||<span data-ttu-id="1f699-p150">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1f699-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1f699-885">**OU**</span><span class="sxs-lookup"><span data-stu-id="1f699-885">**OR**</span></span><br/><span data-ttu-id="1f699-p151">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="1f699-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1f699-888">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-888">String</span></span>|<span data-ttu-id="1f699-889">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-889">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-p152">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1f699-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1f699-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1f699-893">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-893">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-894">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1f699-895">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-895">String</span></span>||<span data-ttu-id="1f699-p153">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1f699-898">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-898">String</span></span>||<span data-ttu-id="1f699-899">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="1f699-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1f699-900">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-900">String</span></span>||<span data-ttu-id="1f699-p154">Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1f699-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1f699-903">Booleano</span><span class="sxs-lookup"><span data-stu-id="1f699-903">Boolean</span></span>||<span data-ttu-id="1f699-p155">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="1f699-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1f699-906">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-906">String</span></span>||<span data-ttu-id="1f699-p156">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1f699-910">função</span><span class="sxs-lookup"><span data-stu-id="1f699-910">function</span></span>|<span data-ttu-id="1f699-911">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-911">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-912">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-913">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-913">Requirements</span></span>

|<span data-ttu-id="1f699-914">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-914">Requirement</span></span>|<span data-ttu-id="1f699-915">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-916">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-916">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-917">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-917">1.0</span></span>|
|[<span data-ttu-id="1f699-918">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-919">ReadItem</span></span>|
|[<span data-ttu-id="1f699-920">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-921">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1f699-922">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1f699-922">Examples</span></span>

<span data-ttu-id="1f699-923">O código a seguir passa uma sequência de caracteres para a função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1f699-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1f699-924">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="1f699-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1f699-925">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="1f699-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1f699-926">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="1f699-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1f699-927">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1f699-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1f699-928">Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1f699-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1f699-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1f699-930">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1f699-930">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-931">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-931">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-932">Requirements</span></span>

|<span data-ttu-id="1f699-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-933">Requirement</span></span>|<span data-ttu-id="1f699-934">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-935">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-936">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-936">1.0</span></span>|
|[<span data-ttu-id="1f699-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-938">ReadItem</span></span>|
|[<span data-ttu-id="1f699-939">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-940">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-941">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-941">Returns:</span></span>

<span data-ttu-id="1f699-942">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1f699-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1f699-943">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-943">Example</span></span>

<span data-ttu-id="1f699-944">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-944">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1f699-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1f699-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1f699-946">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1f699-946">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-947">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-947">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-948">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-948">Parameters:</span></span>

|<span data-ttu-id="1f699-949">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-949">Name</span></span>|<span data-ttu-id="1f699-950">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-950">Type</span></span>|<span data-ttu-id="1f699-951">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="1f699-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1f699-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="1f699-953">Um dos valores da enumeração EntityType.</span><span class="sxs-lookup"><span data-stu-id="1f699-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-954">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-954">Requirements</span></span>

|<span data-ttu-id="1f699-955">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-955">Requirement</span></span>|<span data-ttu-id="1f699-956">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-957">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-958">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-958">1.0</span></span>|
|[<span data-ttu-id="1f699-959">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-960">Restrito</span><span class="sxs-lookup"><span data-stu-id="1f699-960">Restricted</span></span>|
|[<span data-ttu-id="1f699-961">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-962">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-963">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-963">Returns:</span></span>

<span data-ttu-id="1f699-964">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="1f699-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1f699-965">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1f699-965">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="1f699-966">Caso contrário, o tipo dos objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1f699-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1f699-967">Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="1f699-968">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="1f699-968">Value of `entityType`</span></span>|<span data-ttu-id="1f699-969">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="1f699-969">Type of objects in returned array</span></span>|<span data-ttu-id="1f699-970">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="1f699-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="1f699-971">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-971">String</span></span>|<span data-ttu-id="1f699-972">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1f699-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="1f699-973">Contact</span><span class="sxs-lookup"><span data-stu-id="1f699-973">Contact</span></span>|<span data-ttu-id="1f699-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1f699-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="1f699-975">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-975">String</span></span>|<span data-ttu-id="1f699-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1f699-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="1f699-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1f699-977">MeetingSuggestion</span></span>|<span data-ttu-id="1f699-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1f699-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="1f699-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1f699-979">PhoneNumber</span></span>|<span data-ttu-id="1f699-980">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1f699-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="1f699-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1f699-981">TaskSuggestion</span></span>|<span data-ttu-id="1f699-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1f699-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="1f699-983">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-983">String</span></span>|<span data-ttu-id="1f699-984">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="1f699-984">**Restricted**</span></span>|

<span data-ttu-id="1f699-985">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1f699-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1f699-986">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-986">Example</span></span>

<span data-ttu-id="1f699-987">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa os endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="1f699-987">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1f699-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1f699-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1f699-989">Retorna entidades conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1f699-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-990">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-990">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-991">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor especificado no elemento `FilterName` .</span><span class="sxs-lookup"><span data-stu-id="1f699-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-992">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-992">Parameters:</span></span>

|<span data-ttu-id="1f699-993">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-993">Name</span></span>|<span data-ttu-id="1f699-994">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-994">Type</span></span>|<span data-ttu-id="1f699-995">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1f699-996">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-996">String</span></span>|<span data-ttu-id="1f699-997">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="1f699-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-998">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-998">Requirements</span></span>

|<span data-ttu-id="1f699-999">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-999">Requirement</span></span>|<span data-ttu-id="1f699-1000">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1001">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1001">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-1002">1.0</span></span>|
|[<span data-ttu-id="1f699-1003">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1004">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1005">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1006">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1007">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1007">Returns:</span></span>

<span data-ttu-id="1f699-p158">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="1f699-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1f699-1010">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1f699-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1f699-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1f699-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1f699-1012">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1f699-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1013">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-1013">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-p159">O método `getRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="1f699-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1f699-1017">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="1f699-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1f699-1018">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1f699-p160">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="1f699-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-1022">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1022">Requirements</span></span>

|<span data-ttu-id="1f699-1023">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1023">Requirement</span></span>|<span data-ttu-id="1f699-1024">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1025">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1025">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-1026">1.0</span></span>|
|[<span data-ttu-id="1f699-1027">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1028">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1029">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1030">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1031">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1031">Returns:</span></span>

<span data-ttu-id="1f699-p161">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="1f699-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1f699-1034">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="1f699-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1f699-1035">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1f699-1036">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1036">Example</span></span>

<span data-ttu-id="1f699-1037">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="1f699-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1f699-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1f699-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1f699-1039">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1f699-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1040">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-1041">O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="1f699-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1f699-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="1f699-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1044">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1044">Parameters:</span></span>

|<span data-ttu-id="1f699-1045">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1045">Name</span></span>|<span data-ttu-id="1f699-1046">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1046">Type</span></span>|<span data-ttu-id="1f699-1047">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1f699-1048">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-1048">String</span></span>|<span data-ttu-id="1f699-1049">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="1f699-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1050">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1050">Requirements</span></span>

|<span data-ttu-id="1f699-1051">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1051">Requirement</span></span>|<span data-ttu-id="1f699-1052">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1053">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-1054">1.0</span></span>|
|[<span data-ttu-id="1f699-1055">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1056">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1057">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1058">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1059">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1059">Returns:</span></span>

<span data-ttu-id="1f699-1060">Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="1f699-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1f699-1061">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="1f699-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1f699-1062">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1f699-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1f699-1063">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1f699-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1f699-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1f699-1065">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1f699-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1f699-p163">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="1f699-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1068">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1068">Parameters:</span></span>

|<span data-ttu-id="1f699-1069">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1069">Name</span></span>|<span data-ttu-id="1f699-1070">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1070">Type</span></span>|<span data-ttu-id="1f699-1071">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1071">Attributes</span></span>|<span data-ttu-id="1f699-1072">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="1f699-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1f699-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1f699-p164">Solicita um formato para os dados. Se for Text, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="1f699-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="1f699-1077">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1077">Object</span></span>|<span data-ttu-id="1f699-1078">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1079">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-1080">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1080">Object</span></span>|<span data-ttu-id="1f699-1081">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1082">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1f699-1083">function</span><span class="sxs-lookup"><span data-stu-id="1f699-1083">function</span></span>||<span data-ttu-id="1f699-1084">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1f699-1085">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1f699-1086">Para acessar a propriedade de origem de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1086">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1087">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1087">Requirements</span></span>

|<span data-ttu-id="1f699-1088">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1088">Requirement</span></span>|<span data-ttu-id="1f699-1089">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1090">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="1f699-1091">1.2</span></span>|
|[<span data-ttu-id="1f699-1092">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-1094">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1095">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1096">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1096">Returns:</span></span>

<span data-ttu-id="1f699-1097">Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1f699-1098">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="1f699-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1f699-1099">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1f699-1100">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1100">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1f699-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1f699-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1f699-p166">Obtém as entidades encontradas em uma correspondência destacada que um usuário selecionou. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="1f699-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1104">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-1104">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-1105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1105">Requirements</span></span>

|<span data-ttu-id="1f699-1106">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1106">Requirement</span></span>|<span data-ttu-id="1f699-1107">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="1f699-1109">-16</span></span>|
|[<span data-ttu-id="1f699-1110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1111">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1113">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1114">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1114">Returns:</span></span>

<span data-ttu-id="1f699-1115">Tipo: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1f699-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1f699-1116">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1116">Example</span></span>

<span data-ttu-id="1f699-1117">O exemplo a seguir acessa as entidades de endereços na correspondência destacada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="1f699-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1f699-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1f699-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1f699-p167">Retorna valores do tipo sequência de caracteres em uma correspondência destacada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="1f699-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1121">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1f699-1121">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1f699-p168">O método `getSelectedRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="1f699-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1f699-1125">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="1f699-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1f699-1126">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1f699-p169">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="1f699-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f699-1130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1130">Requirements</span></span>

|<span data-ttu-id="1f699-1131">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1131">Requirement</span></span>|<span data-ttu-id="1f699-1132">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1133">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="1f699-1134">-16</span></span>|
|[<span data-ttu-id="1f699-1135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1136">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1138">Leitura</span><span class="sxs-lookup"><span data-stu-id="1f699-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1f699-1139">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1f699-1139">Returns:</span></span>

<span data-ttu-id="1f699-p170">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="1f699-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1f699-1142">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1142">Example</span></span>

<span data-ttu-id="1f699-1143">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="1f699-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1f699-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1f699-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1f699-1145">Carrega de forma assíncrona as propriedades personalizadas desse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="1f699-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1f699-p171">As propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item e o suplemento atuais. As propriedades personalizadas não são criptografadas no item, portanto, isto não deve ser usado como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="1f699-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1149">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1149">Parameters:</span></span>

|<span data-ttu-id="1f699-1150">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1150">Name</span></span>|<span data-ttu-id="1f699-1151">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1151">Type</span></span>|<span data-ttu-id="1f699-1152">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1152">Attributes</span></span>|<span data-ttu-id="1f699-1153">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="1f699-1154">function</span><span class="sxs-lookup"><span data-stu-id="1f699-1154">function</span></span>||<span data-ttu-id="1f699-1155">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1f699-1156">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1f699-1157">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas no servidor.</span><span class="sxs-lookup"><span data-stu-id="1f699-1157">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="1f699-1158">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1158">Object</span></span>|<span data-ttu-id="1f699-1159">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1160">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1160">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="1f699-1161">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1162">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1162">Requirements</span></span>

|<span data-ttu-id="1f699-1163">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1163">Requirement</span></span>|<span data-ttu-id="1f699-1164">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1165">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="1f699-1166">1.0</span></span>|
|[<span data-ttu-id="1f699-1167">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1168">ReadItem</span></span>|
|[<span data-ttu-id="1f699-1169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1170">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-1171">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1171">Example</span></span>

<span data-ttu-id="1f699-p174">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1f699-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1f699-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1f699-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1f699-1176">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1f699-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1f699-p175">O método `removeAttachmentAsync` remove do item o anexo com o identificador especificado. Conforme as práticas recomendadas, você deve usar o identificador do anexo para remover o anexo apenas se o mesmo aplicativo de email tiver inserido o anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador de anexos é válido somente dentro da mesma sessão. Uma sessão é considerada encerrada quando o usuário fecha o aplicativo, ou se o usuário começa a escrever um email em um formulário embutido e, em seguida, abre o mesmo formulário em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="1f699-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1181">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1181">Parameters:</span></span>

|<span data-ttu-id="1f699-1182">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1182">Name</span></span>|<span data-ttu-id="1f699-1183">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1183">Type</span></span>|<span data-ttu-id="1f699-1184">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1184">Attributes</span></span>|<span data-ttu-id="1f699-1185">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="1f699-1186">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-1186">String</span></span>||<span data-ttu-id="1f699-p176">O identificador do anexo a ser removido. O comprimento máximo da sequência de caracteres é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1f699-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="1f699-1189">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1189">Object</span></span>|<span data-ttu-id="1f699-1190">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1191">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-1192">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1192">Object</span></span>|<span data-ttu-id="1f699-1193">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1194">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1f699-1195">function</span><span class="sxs-lookup"><span data-stu-id="1f699-1195">function</span></span>|<span data-ttu-id="1f699-1196">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1197">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1f699-1198">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="1f699-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1f699-1199">Erros</span><span class="sxs-lookup"><span data-stu-id="1f699-1199">Errors</span></span>

|<span data-ttu-id="1f699-1200">Código de erro</span><span class="sxs-lookup"><span data-stu-id="1f699-1200">Error code</span></span>|<span data-ttu-id="1f699-1201">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="1f699-1202">O identificador do anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="1f699-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1203">Requirements</span></span>

|<span data-ttu-id="1f699-1204">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1204">Requirement</span></span>|<span data-ttu-id="1f699-1205">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1206">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="1f699-1207">1.1</span></span>|
|[<span data-ttu-id="1f699-1208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-1210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1211">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-1212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1212">Example</span></span>

<span data-ttu-id="1f699-1213">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="1f699-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1f699-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1f699-1214">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1f699-1215">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="1f699-1215">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1f699-1216">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged` , `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="1f699-1216">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1217">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1217">Parameters:</span></span>

| <span data-ttu-id="1f699-1218">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1218">Name</span></span> | <span data-ttu-id="1f699-1219">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1219">Type</span></span> | <span data-ttu-id="1f699-1220">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1220">Attributes</span></span> | <span data-ttu-id="1f699-1221">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1f699-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1f699-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1f699-1223">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="1f699-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1f699-1224">Função</span><span class="sxs-lookup"><span data-stu-id="1f699-1224">Function</span></span> || <span data-ttu-id="1f699-p177">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1f699-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1f699-1228">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1228">Object</span></span> | <span data-ttu-id="1f699-1229">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="1f699-1230">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1f699-1231">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1231">Object</span></span> | <span data-ttu-id="1f699-1232">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="1f699-1233">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1f699-1234">function</span><span class="sxs-lookup"><span data-stu-id="1f699-1234">function</span></span>| <span data-ttu-id="1f699-1235">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1236">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1237">Requirements</span></span>

|<span data-ttu-id="1f699-1238">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1238">Requirement</span></span>| <span data-ttu-id="1f699-1239">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1240">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f699-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="1f699-1241">-17</span></span> |
|[<span data-ttu-id="1f699-1242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f699-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1243">ReadItem</span></span> |
|[<span data-ttu-id="1f699-1244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f699-1245">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f699-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="1f699-1246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1f699-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1f699-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="1f699-1248">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1f699-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="1f699-p178">Quando chamado, este método salva a mensagem atual como um rascunho e retorna o identificador do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook em modo de cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="1f699-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1252">Se o seu suplemento chamar `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver em modo de cache, poderá levar algum tempo antes do item realmente ser sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="1f699-1252">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="1f699-1253">Até que o item seja sincronizado, o uso de `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="1f699-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1f699-p180">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo de redação, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="1f699-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1f699-1257">Os seguintes clientes possuem um comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="1f699-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1f699-1258">O Outlook para Mac não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="1f699-1258">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="1f699-1259">Chamar `saveAsync` em uma reunião no Outlook para Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="1f699-1259">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1f699-1260">O Outlook na Web sempre enviará um convite ou atualização quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="1f699-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1261">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1261">Parameters:</span></span>

|<span data-ttu-id="1f699-1262">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1262">Name</span></span>|<span data-ttu-id="1f699-1263">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1263">Type</span></span>|<span data-ttu-id="1f699-1264">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1264">Attributes</span></span>|<span data-ttu-id="1f699-1265">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1f699-1266">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1266">Object</span></span>|<span data-ttu-id="1f699-1267">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1268">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-1269">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1269">Object</span></span>|<span data-ttu-id="1f699-1270">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1271">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1f699-1272">function</span><span class="sxs-lookup"><span data-stu-id="1f699-1272">function</span></span>||<span data-ttu-id="1f699-1273">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1f699-1274">Em caso de sucesso, o identificador do item será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1f699-1274">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1275">Requirements</span></span>

|<span data-ttu-id="1f699-1276">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1276">Requirement</span></span>|<span data-ttu-id="1f699-1277">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1278">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="1f699-1279">1.3</span></span>|
|[<span data-ttu-id="1f699-1280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-1282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1283">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1f699-1284">Exemplos</span><span class="sxs-lookup"><span data-stu-id="1f699-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="1f699-p182">A seguir apresentamos um exemplo do parâmetro `result` passado para a função de retorno de chamada. A propriedade `value` contém o ID do item.</span><span class="sxs-lookup"><span data-stu-id="1f699-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1f699-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1f699-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1f699-1288">Insere dados no corpo ou no assunto de uma mensagem de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1f699-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1f699-p183">O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou no assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou do assunto, um erro será retornado. Após a inserção, o cursor será posicionado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="1f699-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1f699-1292">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1f699-1292">Parameters:</span></span>

|<span data-ttu-id="1f699-1293">Nome</span><span class="sxs-lookup"><span data-stu-id="1f699-1293">Name</span></span>|<span data-ttu-id="1f699-1294">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f699-1294">Type</span></span>|<span data-ttu-id="1f699-1295">Atributos</span><span class="sxs-lookup"><span data-stu-id="1f699-1295">Attributes</span></span>|<span data-ttu-id="1f699-1296">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f699-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="1f699-1297">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f699-1297">String</span></span>||<span data-ttu-id="1f699-p184">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="1f699-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="1f699-1301">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1301">Object</span></span>|<span data-ttu-id="1f699-1302">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1303">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f699-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1f699-1304">Objeto</span><span class="sxs-lookup"><span data-stu-id="1f699-1304">Object</span></span>|<span data-ttu-id="1f699-1305">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-1306">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1f699-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1f699-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1f699-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1f699-1308">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1f699-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="1f699-p185">Se for `text` , o estilo atual será aplicado no Outlook Web App e no Outlook. Se o campo for um editor HTML, somente os dados de texto serão inseridos, mesmo que os dados estejam em HTML.</span><span class="sxs-lookup"><span data-stu-id="1f699-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1f699-p186">Se for `html` e o campo for compatível com HTML (e o assunto não), o estilo atual será aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, um erro `InvalidDataFormat` será retornado.</span><span class="sxs-lookup"><span data-stu-id="1f699-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1f699-1313">Se `coercionType` não estiver definido, o resultado dependerá do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="1f699-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="1f699-1314">função</span><span class="sxs-lookup"><span data-stu-id="1f699-1314">function</span></span>||<span data-ttu-id="1f699-1315">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1f699-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f699-1316">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f699-1316">Requirements</span></span>

|<span data-ttu-id="1f699-1317">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f699-1317">Requirement</span></span>|<span data-ttu-id="1f699-1318">Valor</span><span class="sxs-lookup"><span data-stu-id="1f699-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f699-1319">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f699-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1f699-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="1f699-1320">1.2</span></span>|
|[<span data-ttu-id="1f699-1321">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f699-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1f699-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1f699-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="1f699-1323">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1f699-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1f699-1324">Redação</span><span class="sxs-lookup"><span data-stu-id="1f699-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1f699-1325">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f699-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```