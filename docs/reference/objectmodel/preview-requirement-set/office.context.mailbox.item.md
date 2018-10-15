
# <a name="item"></a><span data-ttu-id="f6309-101">item</span><span class="sxs-lookup"><span data-stu-id="f6309-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f6309-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f6309-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f6309-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f6309-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-105">Requirements</span></span>

|<span data-ttu-id="f6309-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-106">Requirement</span></span>|<span data-ttu-id="f6309-107">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-109">1.0</span></span>|
|[<span data-ttu-id="f6309-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="f6309-111">Restricted</span></span>|
|[<span data-ttu-id="f6309-112">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-113">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f6309-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="f6309-114">Members and methods</span></span>

| <span data-ttu-id="f6309-115">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-115">Member</span></span> | <span data-ttu-id="f6309-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f6309-117">attachments</span><span class="sxs-lookup"><span data-stu-id="f6309-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="f6309-118">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-118">Member</span></span> |
| [<span data-ttu-id="f6309-119">bcc</span><span class="sxs-lookup"><span data-stu-id="f6309-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="f6309-120">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-120">Member</span></span> |
| [<span data-ttu-id="f6309-121">body</span><span class="sxs-lookup"><span data-stu-id="f6309-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="f6309-122">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-122">Member</span></span> |
| [<span data-ttu-id="f6309-123">cc</span><span class="sxs-lookup"><span data-stu-id="f6309-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="f6309-124">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-124">Member</span></span> |
| [<span data-ttu-id="f6309-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="f6309-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f6309-126">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-126">Member</span></span> |
| [<span data-ttu-id="f6309-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f6309-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f6309-128">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-128">Member</span></span> |
| [<span data-ttu-id="f6309-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f6309-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f6309-130">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-130">Member</span></span> |
| [<span data-ttu-id="f6309-131">end</span><span class="sxs-lookup"><span data-stu-id="f6309-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="f6309-132">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-132">Member</span></span> |
| [<span data-ttu-id="f6309-133">from</span><span class="sxs-lookup"><span data-stu-id="f6309-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="f6309-134">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-134">Member</span></span> |
| [<span data-ttu-id="f6309-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f6309-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f6309-136">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-136">Member</span></span> |
| [<span data-ttu-id="f6309-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="f6309-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f6309-138">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-138">Member</span></span> |
| [<span data-ttu-id="f6309-139">itemId</span><span class="sxs-lookup"><span data-stu-id="f6309-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f6309-140">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-140">Member</span></span> |
| [<span data-ttu-id="f6309-141">itemType</span><span class="sxs-lookup"><span data-stu-id="f6309-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="f6309-142">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-142">Member</span></span> |
| [<span data-ttu-id="f6309-143">location</span><span class="sxs-lookup"><span data-stu-id="f6309-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="f6309-144">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-144">Member</span></span> |
| [<span data-ttu-id="f6309-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f6309-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f6309-146">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-146">Member</span></span> |
| [<span data-ttu-id="f6309-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f6309-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="f6309-148">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-148">Member</span></span> |
| [<span data-ttu-id="f6309-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f6309-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="f6309-150">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-150">Member</span></span> |
| [<span data-ttu-id="f6309-151">organizer</span><span class="sxs-lookup"><span data-stu-id="f6309-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="f6309-152">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-152">Member</span></span> |
| [<span data-ttu-id="f6309-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="f6309-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="f6309-154">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-154">Member</span></span> |
| [<span data-ttu-id="f6309-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f6309-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="f6309-156">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-156">Member</span></span> |
| [<span data-ttu-id="f6309-157">sender</span><span class="sxs-lookup"><span data-stu-id="f6309-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="f6309-158">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-158">Member</span></span> |
| [<span data-ttu-id="f6309-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="f6309-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="f6309-160">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-160">Member</span></span> |
| [<span data-ttu-id="f6309-161">start</span><span class="sxs-lookup"><span data-stu-id="f6309-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="f6309-162">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-162">Member</span></span> |
| [<span data-ttu-id="f6309-163">subject</span><span class="sxs-lookup"><span data-stu-id="f6309-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="f6309-164">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-164">Member</span></span> |
| [<span data-ttu-id="f6309-165">to</span><span class="sxs-lookup"><span data-stu-id="f6309-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="f6309-166">Membro</span><span class="sxs-lookup"><span data-stu-id="f6309-166">Member</span></span> |
| [<span data-ttu-id="f6309-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f6309-168">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-168">Method</span></span> |
| [<span data-ttu-id="f6309-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f6309-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="f6309-170">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-170">Method</span></span> |
| [<span data-ttu-id="f6309-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f6309-172">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-172">Method</span></span> |
| [<span data-ttu-id="f6309-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f6309-174">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-174">Method</span></span> |
| [<span data-ttu-id="f6309-175">close</span><span class="sxs-lookup"><span data-stu-id="f6309-175">close</span></span>](#close) | <span data-ttu-id="f6309-176">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-176">Method</span></span> |
| [<span data-ttu-id="f6309-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f6309-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="f6309-178">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-178">Method</span></span> |
| [<span data-ttu-id="f6309-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f6309-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="f6309-180">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-180">Method</span></span> |
| [<span data-ttu-id="f6309-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="f6309-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="f6309-182">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-182">Method</span></span> |
| [<span data-ttu-id="f6309-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f6309-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="f6309-184">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-184">Method</span></span> |
| [<span data-ttu-id="f6309-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f6309-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="f6309-186">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-186">Method</span></span> |
| [<span data-ttu-id="f6309-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="f6309-188">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-188">Method</span></span> |
| [<span data-ttu-id="f6309-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f6309-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f6309-190">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-190">Method</span></span> |
| [<span data-ttu-id="f6309-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f6309-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f6309-192">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-192">Method</span></span> |
| [<span data-ttu-id="f6309-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f6309-194">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-194">Method</span></span> |
| [<span data-ttu-id="f6309-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="f6309-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="f6309-196">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-196">Method</span></span> |
| [<span data-ttu-id="f6309-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f6309-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f6309-198">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-198">Method</span></span> |
| [<span data-ttu-id="f6309-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="f6309-200">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-200">Method</span></span> |
| [<span data-ttu-id="f6309-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f6309-202">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-202">Method</span></span> |
| [<span data-ttu-id="f6309-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f6309-204">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-204">Method</span></span> |
| [<span data-ttu-id="f6309-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f6309-206">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-206">Method</span></span> |
| [<span data-ttu-id="f6309-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f6309-208">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-208">Method</span></span> |
| [<span data-ttu-id="f6309-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f6309-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f6309-210">Método</span><span class="sxs-lookup"><span data-stu-id="f6309-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f6309-211">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-211">Example</span></span>

<span data-ttu-id="f6309-212">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6309-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f6309-213">Membros</span><span class="sxs-lookup"><span data-stu-id="f6309-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="f6309-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f6309-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="f6309-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-217">Certos tipos de arquivos são bloqueados pelo Outlook devido a potenciais problemas de segurança e portanto não são retornados.</span><span class="sxs-lookup"><span data-stu-id="f6309-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f6309-218">Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f6309-218">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-219">Type:</span></span>

*   <span data-ttu-id="f6309-220">Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f6309-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-221">Requirements</span></span>

|<span data-ttu-id="f6309-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-222">Requirement</span></span>|<span data-ttu-id="f6309-223">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-224">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-225">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-225">1.0</span></span>|
|[<span data-ttu-id="f6309-226">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-227">ReadItem</span></span>|
|[<span data-ttu-id="f6309-228">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-229">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-230">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-230">Example</span></span>

<span data-ttu-id="f6309-231">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f6309-232">cco:[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f6309-233">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f6309-234">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f6309-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-235">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-235">Type:</span></span>

*   [<span data-ttu-id="f6309-236">Destinatários</span><span class="sxs-lookup"><span data-stu-id="f6309-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f6309-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-237">Requirements</span></span>

|<span data-ttu-id="f6309-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-238">Requirement</span></span>|<span data-ttu-id="f6309-239">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-240">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-241">1.1</span><span class="sxs-lookup"><span data-stu-id="f6309-241">1.1</span></span>|
|[<span data-ttu-id="f6309-242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-243">ReadItem</span></span>|
|[<span data-ttu-id="f6309-244">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-245">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="f6309-247">corpo:[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="f6309-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="f6309-248">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="f6309-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-249">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-249">Type:</span></span>

*   [<span data-ttu-id="f6309-250">Body</span><span class="sxs-lookup"><span data-stu-id="f6309-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="f6309-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-251">Requirements</span></span>

|<span data-ttu-id="f6309-252">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-252">Requirement</span></span>|<span data-ttu-id="f6309-253">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-254">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-255">1.1</span><span class="sxs-lookup"><span data-stu-id="f6309-255">1.1</span></span>|
|[<span data-ttu-id="f6309-256">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-257">ReadItem</span></span>|
|[<span data-ttu-id="f6309-258">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-259">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f6309-260">cc: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f6309-261">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f6309-262">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-263">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-263">Read mode</span></span>

<span data-ttu-id="f6309-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f6309-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-266">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-266">Compose mode</span></span>

<span data-ttu-id="f6309-267">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-267">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-268">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-268">Type:</span></span>

*   <span data-ttu-id="f6309-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-270">Requirements</span></span>

|<span data-ttu-id="f6309-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-271">Requirement</span></span>|<span data-ttu-id="f6309-272">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-273">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-274">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-274">1.0</span></span>|
|[<span data-ttu-id="f6309-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-276">ReadItem</span></span>|
|[<span data-ttu-id="f6309-277">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-278">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f6309-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f6309-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="f6309-281">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="f6309-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f6309-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviá-la, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não será mais aplicável.</span><span class="sxs-lookup"><span data-stu-id="f6309-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f6309-p108">Para um novo item em um formulário de redação, o valor dessa propriedade é nulo. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="f6309-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-286">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-286">Type:</span></span>

*   <span data-ttu-id="f6309-287">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-288">Requirements</span></span>

|<span data-ttu-id="f6309-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-289">Requirement</span></span>|<span data-ttu-id="f6309-290">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-291">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-292">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-292">1.0</span></span>|
|[<span data-ttu-id="f6309-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-294">ReadItem</span></span>|
|[<span data-ttu-id="f6309-295">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-296">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f6309-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="f6309-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="f6309-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-300">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-300">Type:</span></span>

*   <span data-ttu-id="f6309-301">Data</span><span class="sxs-lookup"><span data-stu-id="f6309-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-302">Requirements</span></span>

|<span data-ttu-id="f6309-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-303">Requirement</span></span>|<span data-ttu-id="f6309-304">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-305">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-306">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-306">1.0</span></span>|
|[<span data-ttu-id="f6309-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-308">ReadItem</span></span>|
|[<span data-ttu-id="f6309-309">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-310">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-311">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f6309-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="f6309-312">dateTimeModified :Date</span></span>

<span data-ttu-id="f6309-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-315">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-315">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-316">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-316">Type:</span></span>

*   <span data-ttu-id="f6309-317">Data</span><span class="sxs-lookup"><span data-stu-id="f6309-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-318">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-318">Requirements</span></span>

|<span data-ttu-id="f6309-319">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-319">Requirement</span></span>|<span data-ttu-id="f6309-320">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-321">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-321">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-322">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-322">1.0</span></span>|
|[<span data-ttu-id="f6309-323">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-324">ReadItem</span></span>|
|[<span data-ttu-id="f6309-325">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-326">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-327">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f6309-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f6309-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f6309-329">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="f6309-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f6309-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade para a data e hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f6309-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-332">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-332">Read mode</span></span>

<span data-ttu-id="f6309-333">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f6309-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-334">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-334">Compose mode</span></span>

<span data-ttu-id="f6309-335">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f6309-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f6309-336">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC do servidor.</span><span class="sxs-lookup"><span data-stu-id="f6309-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-337">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-337">Type:</span></span>

*   <span data-ttu-id="f6309-338">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f6309-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-339">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-339">Requirements</span></span>

|<span data-ttu-id="f6309-340">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-340">Requirement</span></span>|<span data-ttu-id="f6309-341">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-342">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-343">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-343">1.0</span></span>|
|[<span data-ttu-id="f6309-344">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-345">ReadItem</span></span>|
|[<span data-ttu-id="f6309-346">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-347">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-348">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-348">Example</span></span>

<span data-ttu-id="f6309-349">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f6309-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="f6309-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f6309-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="f6309-351">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="f6309-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="f6309-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-354">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f6309-354">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-355">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-355">Read mode</span></span>

<span data-ttu-id="f6309-356">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="f6309-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="f6309-357">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-357">Compose mode</span></span>

<span data-ttu-id="f6309-358">A propriedade `from` retornará um objeto `From` que fornece um método para obter o valor de from.</span><span class="sxs-lookup"><span data-stu-id="f6309-358">Added From: Adds a new object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f6309-359">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-359">Type:</span></span>

*   <span data-ttu-id="f6309-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f6309-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-361">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-361">Requirements</span></span>

|<span data-ttu-id="f6309-362">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f6309-363">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-364">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-364">1.0</span></span>|<span data-ttu-id="f6309-365">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-365">-17</span></span>|
|[<span data-ttu-id="f6309-366">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-367">ReadItem</span></span>|<span data-ttu-id="f6309-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-369">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-370">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-370">Read</span></span>|<span data-ttu-id="f6309-371">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f6309-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f6309-372">internetMessageId :String</span></span>

<span data-ttu-id="f6309-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-375">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-375">Type:</span></span>

*   <span data-ttu-id="f6309-376">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-377">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-377">Requirements</span></span>

|<span data-ttu-id="f6309-378">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-378">Requirement</span></span>|<span data-ttu-id="f6309-379">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-380">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-381">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-381">1.0</span></span>|
|[<span data-ttu-id="f6309-382">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-383">ReadItem</span></span>|
|[<span data-ttu-id="f6309-384">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-385">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-386">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f6309-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f6309-387">itemClass :String</span></span>

<span data-ttu-id="f6309-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f6309-p115">A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir estão as classes de mensagem padrão para itens de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="f6309-392">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-392">Type</span></span>|<span data-ttu-id="f6309-393">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-393">Description</span></span>|<span data-ttu-id="f6309-394">classe do item</span><span class="sxs-lookup"><span data-stu-id="f6309-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="f6309-395">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="f6309-395">Appointment items</span></span>|<span data-ttu-id="f6309-396">São itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="f6309-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="f6309-397">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="f6309-397">Message items</span></span>|<span data-ttu-id="f6309-398">Incluem mensagens de e-mail que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagens base.</span><span class="sxs-lookup"><span data-stu-id="f6309-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="f6309-399">Você pode criar classes de mensagens personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="f6309-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-400">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-400">Type:</span></span>

*   <span data-ttu-id="f6309-401">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-402">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-402">Requirements</span></span>

|<span data-ttu-id="f6309-403">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-403">Requirement</span></span>|<span data-ttu-id="f6309-404">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-405">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-406">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-406">1.0</span></span>|
|[<span data-ttu-id="f6309-407">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-408">ReadItem</span></span>|
|[<span data-ttu-id="f6309-409">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-410">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-411">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f6309-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="f6309-412">(nullable) itemId :String</span></span>

<span data-ttu-id="f6309-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-415">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="f6309-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f6309-416">A propriedade `itemId` não é idêntica ao Entry ID do Outlook ou ao ID usado pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6309-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f6309-417">Antes de fazer chamadas à API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f6309-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f6309-418">Para mais informações, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f6309-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f6309-p118">A propriedade `itemId` não está disponível no modo de redação. Se  um identificador de item for requerido, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no repositório, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-421">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-421">Type:</span></span>

*   <span data-ttu-id="f6309-422">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-423">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-423">Requirements</span></span>

|<span data-ttu-id="f6309-424">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-424">Requirement</span></span>|<span data-ttu-id="f6309-425">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-426">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-427">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-427">1.0</span></span>|
|[<span data-ttu-id="f6309-428">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-429">ReadItem</span></span>|
|[<span data-ttu-id="f6309-430">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-431">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-432">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-432">Example</span></span>

<span data-ttu-id="f6309-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item a partir do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f6309-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="f6309-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f6309-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f6309-436">Obtém o tipo de item que uma instância representa.</span><span class="sxs-lookup"><span data-stu-id="f6309-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f6309-437">A propriedade `itemType` retorna um dos valores da enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-438">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-438">Type:</span></span>

*   [<span data-ttu-id="f6309-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f6309-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f6309-440">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-440">Requirements</span></span>

|<span data-ttu-id="f6309-441">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-441">Requirement</span></span>|<span data-ttu-id="f6309-442">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-443">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-444">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-444">1.0</span></span>|
|[<span data-ttu-id="f6309-445">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-446">ReadItem</span></span>|
|[<span data-ttu-id="f6309-447">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-448">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-449">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="f6309-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f6309-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="f6309-451">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-452">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-452">Read mode</span></span>

<span data-ttu-id="f6309-453">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-454">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-454">Compose mode</span></span>

<span data-ttu-id="f6309-455">A propriedade `location` retorna um objeto `Location` que fornece métodos para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-456">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-456">Type:</span></span>

*   <span data-ttu-id="f6309-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f6309-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-458">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-458">Requirements</span></span>

|<span data-ttu-id="f6309-459">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-459">Requirement</span></span>|<span data-ttu-id="f6309-460">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-461">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-462">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-462">1.0</span></span>|
|[<span data-ttu-id="f6309-463">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-464">ReadItem</span></span>|
|[<span data-ttu-id="f6309-465">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-466">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-467">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f6309-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f6309-468">normalizedSubject :String</span></span>

<span data-ttu-id="f6309-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f6309-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="f6309-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-473">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-473">Type:</span></span>

*   <span data-ttu-id="f6309-474">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-475">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-475">Requirements</span></span>

|<span data-ttu-id="f6309-476">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-476">Requirement</span></span>|<span data-ttu-id="f6309-477">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-478">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-479">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-479">1.0</span></span>|
|[<span data-ttu-id="f6309-480">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-481">ReadItem</span></span>|
|[<span data-ttu-id="f6309-482">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-483">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-484">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="f6309-485">notificationMessages:[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f6309-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="f6309-486">Obtém as mensagens de notificação para um item.</span><span class="sxs-lookup"><span data-stu-id="f6309-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-487">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-487">Type:</span></span>

*   [<span data-ttu-id="f6309-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f6309-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f6309-489">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-489">Requirements</span></span>

|<span data-ttu-id="f6309-490">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-490">Requirement</span></span>|<span data-ttu-id="f6309-491">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-492">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-493">1.3</span><span class="sxs-lookup"><span data-stu-id="f6309-493">1.3</span></span>|
|[<span data-ttu-id="f6309-494">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-495">ReadItem</span></span>|
|[<span data-ttu-id="f6309-496">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-497">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f6309-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f6309-499">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="f6309-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f6309-500">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-501">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-501">Read mode</span></span>

<span data-ttu-id="f6309-502">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-503">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-503">Compose mode</span></span>

<span data-ttu-id="f6309-504">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-505">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-505">Type:</span></span>

*   <span data-ttu-id="f6309-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-507">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-507">Requirements</span></span>

|<span data-ttu-id="f6309-508">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-508">Requirement</span></span>|<span data-ttu-id="f6309-509">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-510">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-511">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-511">1.0</span></span>|
|[<span data-ttu-id="f6309-512">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-513">ReadItem</span></span>|
|[<span data-ttu-id="f6309-514">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-515">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-516">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="f6309-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f6309-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="f6309-518">Obtém o endereço de email do organizador da reunião para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="f6309-518">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-519">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-519">Read mode</span></span>

<span data-ttu-id="f6309-520">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-521">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-521">Compose mode</span></span>

<span data-ttu-id="f6309-522">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook/office.organizer) que fornece um método para obter o valor do organizador.</span><span class="sxs-lookup"><span data-stu-id="f6309-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-523">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-523">Type:</span></span>

*   <span data-ttu-id="f6309-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f6309-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-525">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-525">Requirements</span></span>

|<span data-ttu-id="f6309-526">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f6309-527">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-527">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-528">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-528">1.0</span></span>|<span data-ttu-id="f6309-529">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-529">-17</span></span>|
|[<span data-ttu-id="f6309-530">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-531">ReadItem</span></span>|<span data-ttu-id="f6309-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-533">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-534">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-534">Read</span></span>|<span data-ttu-id="f6309-535">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-536">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="f6309-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="f6309-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="f6309-538">Obtém ou define o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-538">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="f6309-539">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="f6309-540">Modos de leitura e redação para itens de compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="f6309-541">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="f6309-542">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook/office.recurrence) para solicitações de reuniões ou compromissos recorrentes, se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="f6309-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="f6309-543">`null` é retornada para compromissos únicos e solicitações de reunião de compromissos únicos.</span><span class="sxs-lookup"><span data-stu-id="f6309-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="f6309-544">`undefined` é retornado para mensagens que não fazem solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="f6309-545">Observação: Solicitações de reunião tem um valor `itemClass` de IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="f6309-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="f6309-546">Observação: Se o objeto de recorrência for `null`, isto indica que o objeto é um compromisso único ou uma solicitação de reunião de um compromisso único e NÃO faz parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="f6309-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-547">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-547">Type:</span></span>

* [<span data-ttu-id="f6309-548">Recorrência</span><span class="sxs-lookup"><span data-stu-id="f6309-548">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="f6309-549">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-549">Requirement</span></span>|<span data-ttu-id="f6309-550">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-551">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-552">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-552">-17</span></span>|
|[<span data-ttu-id="f6309-553">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-554">ReadItem</span></span>|
|[<span data-ttu-id="f6309-555">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-556">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f6309-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f6309-558">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="f6309-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f6309-559">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-560">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-560">Read mode</span></span>

<span data-ttu-id="f6309-561">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-562">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-562">Compose mode</span></span>

<span data-ttu-id="f6309-563">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-564">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-564">Type:</span></span>

*   <span data-ttu-id="f6309-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-566">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-566">Requirements</span></span>

|<span data-ttu-id="f6309-567">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-567">Requirement</span></span>|<span data-ttu-id="f6309-568">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-569">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-569">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-570">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-570">1.0</span></span>|
|[<span data-ttu-id="f6309-571">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-572">ReadItem</span></span>|
|[<span data-ttu-id="f6309-573">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-574">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-575">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="f6309-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f6309-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="f6309-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f6309-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f6309-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegador, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="f6309-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-581">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f6309-581">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-582">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-582">Type:</span></span>

*   [<span data-ttu-id="f6309-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f6309-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f6309-584">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-584">Requirements</span></span>

|<span data-ttu-id="f6309-585">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-585">Requirement</span></span>|<span data-ttu-id="f6309-586">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-587">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-588">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-588">1.0</span></span>|
|[<span data-ttu-id="f6309-589">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-590">ReadItem</span></span>|
|[<span data-ttu-id="f6309-591">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-592">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-593">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="f6309-594">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="f6309-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="f6309-595">Obtém a identificação da série a que uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="f6309-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="f6309-596">No OWA e no Outlook, o `seriesId` retornará a identificação de serviços Web do Exchange (EWS) do item pai (série) a que este item pertence.</span><span class="sxs-lookup"><span data-stu-id="f6309-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="f6309-597">No entanto, em iOS e Android, o `seriesId` retornará a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="f6309-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-598">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="f6309-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f6309-599">A propriedade `seriesId` não é idêntica às IDs do Outlook usadas pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f6309-599">The `seriesId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f6309-600">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f6309-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f6309-601">Para mais informações, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="f6309-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="f6309-602">A  `seriesId` propriedade  retornará  `null` para itens que não têm itens pai como compromissos, itens de série, ou solicitações de reunião únicos e retorna `undefined` para todos os itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="f6309-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-603">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-603">Type:</span></span>

* <span data-ttu-id="f6309-604">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-605">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-605">Requirements</span></span>

|<span data-ttu-id="f6309-606">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-606">Requirement</span></span>|<span data-ttu-id="f6309-607">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-608">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-609">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-609">-17</span></span>|
|[<span data-ttu-id="f6309-610">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-611">ReadItem</span></span>|
|[<span data-ttu-id="f6309-612">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-613">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-614">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f6309-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f6309-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f6309-616">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="f6309-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f6309-p130">A propriedade `start` é expressa como um valor de data e valor temporal no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f6309-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-619">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-619">Read mode</span></span>

<span data-ttu-id="f6309-620">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f6309-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-621">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-621">Compose mode</span></span>

<span data-ttu-id="f6309-622">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f6309-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f6309-623">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="f6309-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-624">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-624">Type:</span></span>

*   <span data-ttu-id="f6309-625">Data | [Hora](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f6309-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-626">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-626">Requirements</span></span>

|<span data-ttu-id="f6309-627">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-627">Requirement</span></span>|<span data-ttu-id="f6309-628">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-629">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-630">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-630">1.0</span></span>|
|[<span data-ttu-id="f6309-631">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-632">ReadItem</span></span>|
|[<span data-ttu-id="f6309-633">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-634">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-635">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-635">Example</span></span>

<span data-ttu-id="f6309-636">O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f6309-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="f6309-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f6309-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="f6309-638">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="f6309-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f6309-639">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de e-mail.</span><span class="sxs-lookup"><span data-stu-id="f6309-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-640">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-640">Read mode</span></span>

<span data-ttu-id="f6309-p131">A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f6309-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f6309-643">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-643">Compose mode</span></span>

<span data-ttu-id="f6309-644">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="f6309-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f6309-645">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-645">Type:</span></span>

*   <span data-ttu-id="f6309-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f6309-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-647">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-647">Requirements</span></span>

|<span data-ttu-id="f6309-648">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-648">Requirement</span></span>|<span data-ttu-id="f6309-649">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-650">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-651">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-651">1.0</span></span>|
|[<span data-ttu-id="f6309-652">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-653">ReadItem</span></span>|
|[<span data-ttu-id="f6309-654">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-655">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f6309-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f6309-657">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f6309-658">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f6309-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-659">Read mode</span></span>

<span data-ttu-id="f6309-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **To** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f6309-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f6309-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f6309-662">Compose mode</span></span>

<span data-ttu-id="f6309-663">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **To** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-663">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f6309-664">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f6309-664">Type:</span></span>

*   <span data-ttu-id="f6309-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f6309-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-666">Requirements</span></span>

|<span data-ttu-id="f6309-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-667">Requirement</span></span>|<span data-ttu-id="f6309-668">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-669">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-670">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-670">1.0</span></span>|
|[<span data-ttu-id="f6309-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-672">ReadItem</span></span>|
|[<span data-ttu-id="f6309-673">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-674">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-675">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f6309-676">Métodos</span><span class="sxs-lookup"><span data-stu-id="f6309-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f6309-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f6309-678">Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.</span><span class="sxs-lookup"><span data-stu-id="f6309-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f6309-679">O método `addFileAttachmentAsync` carrega o arquivo da URI especificada e o anexa ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="f6309-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f6309-680">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f6309-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-681">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-681">Parameters:</span></span>
|<span data-ttu-id="f6309-682">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-682">Name</span></span>|<span data-ttu-id="f6309-683">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-683">Type</span></span>|<span data-ttu-id="f6309-684">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-684">Attributes</span></span>|<span data-ttu-id="f6309-685">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="f6309-686">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-686">String</span></span>||<span data-ttu-id="f6309-p134">O URI que fornece a localização do arquivo anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f6309-689">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-689">String</span></span>||<span data-ttu-id="f6309-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f6309-692">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-692">Object</span></span>|<span data-ttu-id="f6309-693">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-693">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-694">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-695">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-695">Object</span></span>|<span data-ttu-id="f6309-696">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-696">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-697">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f6309-698">Booleano</span><span class="sxs-lookup"><span data-stu-id="f6309-698">Boolean</span></span>|<span data-ttu-id="f6309-699">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-699">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-700">Se for `true`, indicará que o anexo será embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f6309-701">function</span><span class="sxs-lookup"><span data-stu-id="f6309-701">function</span></span>|<span data-ttu-id="f6309-702">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-702">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-703">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f6309-704">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f6309-705">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f6309-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f6309-706">Erros</span><span class="sxs-lookup"><span data-stu-id="f6309-706">Errors</span></span>

|<span data-ttu-id="f6309-707">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f6309-707">Error code</span></span>|<span data-ttu-id="f6309-708">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f6309-709">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="f6309-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f6309-710">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="f6309-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f6309-711">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-712">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-712">Requirements</span></span>

|<span data-ttu-id="f6309-713">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-713">Requirement</span></span>|<span data-ttu-id="f6309-714">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-715">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-716">1.1</span><span class="sxs-lookup"><span data-stu-id="f6309-716">1.1</span></span>|
|[<span data-ttu-id="f6309-717">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-719">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-720">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f6309-721">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f6309-721">Examples</span></span>

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

<span data-ttu-id="f6309-722">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="f6309-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-723">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f6309-724">Adiciona um arquivo da codificação base64 a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="f6309-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f6309-725">O método  `addFileAttachmentFromBase64Async` carrega o arquivo da codificação base64 e o anexa ao item no formato de redação.</span><span class="sxs-lookup"><span data-stu-id="f6309-725">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="f6309-726">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="f6309-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="f6309-727">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f6309-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-728">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-728">Parameters:</span></span>
|<span data-ttu-id="f6309-729">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-729">Name</span></span>|<span data-ttu-id="f6309-730">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-730">Type</span></span>|<span data-ttu-id="f6309-731">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-731">Attributes</span></span>|<span data-ttu-id="f6309-732">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="f6309-733">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-733">String</span></span>||<span data-ttu-id="f6309-734">O conteúdo codificado em base64 de uma imagem ou um arquivo a ser adicionado a um email ou um evento.</span><span class="sxs-lookup"><span data-stu-id="f6309-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="f6309-735">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-735">String</span></span>||<span data-ttu-id="f6309-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f6309-738">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-738">Object</span></span>|<span data-ttu-id="f6309-739">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-739">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-740">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-741">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-741">Object</span></span>|<span data-ttu-id="f6309-742">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-742">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-743">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f6309-744">Booleano</span><span class="sxs-lookup"><span data-stu-id="f6309-744">Boolean</span></span>|<span data-ttu-id="f6309-745">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-745">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-746">Se for `true`, indicará que o anexo será embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f6309-747">function</span><span class="sxs-lookup"><span data-stu-id="f6309-747">function</span></span>|<span data-ttu-id="f6309-748">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-748">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-749">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f6309-750">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f6309-751">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f6309-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f6309-752">Erros</span><span class="sxs-lookup"><span data-stu-id="f6309-752">Errors</span></span>

|<span data-ttu-id="f6309-753">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f6309-753">Error code</span></span>|<span data-ttu-id="f6309-754">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f6309-755">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="f6309-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f6309-756">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="f6309-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f6309-757">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-758">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-758">Requirements</span></span>

|<span data-ttu-id="f6309-759">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-759">Requirement</span></span>|<span data-ttu-id="f6309-760">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-761">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-762">Visualização</span><span class="sxs-lookup"><span data-stu-id="f6309-762">Preview</span></span>|
|[<span data-ttu-id="f6309-763">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-765">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-766">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f6309-767">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f6309-767">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f6309-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f6309-769">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="f6309-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f6309-770">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged` , `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="f6309-770">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-771">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-771">Parameters:</span></span>

| <span data-ttu-id="f6309-772">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-772">Name</span></span> | <span data-ttu-id="f6309-773">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-773">Type</span></span> | <span data-ttu-id="f6309-774">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-774">Attributes</span></span> | <span data-ttu-id="f6309-775">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f6309-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f6309-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f6309-777">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="f6309-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f6309-778">Função</span><span class="sxs-lookup"><span data-stu-id="f6309-778">Function</span></span> || <span data-ttu-id="f6309-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="f6309-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f6309-782">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-782">Object</span></span> | <span data-ttu-id="f6309-783">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-783">&lt;optional&gt;</span></span> | <span data-ttu-id="f6309-784">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f6309-785">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-785">Object</span></span> | <span data-ttu-id="f6309-786">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-786">&lt;optional&gt;</span></span> | <span data-ttu-id="f6309-787">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f6309-788">function</span><span class="sxs-lookup"><span data-stu-id="f6309-788">function</span></span>| <span data-ttu-id="f6309-789">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-789">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-790">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-791">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-791">Requirements</span></span>

|<span data-ttu-id="f6309-792">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-792">Requirement</span></span>| <span data-ttu-id="f6309-793">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-794">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-794">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6309-795">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-795">-17</span></span> |
|[<span data-ttu-id="f6309-796">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6309-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-797">ReadItem</span></span> |
|[<span data-ttu-id="f6309-798">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6309-799">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f6309-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f6309-801">Adiciona um item do Exchange, como uma mensagem, como um anexo à mensagem ou ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f6309-p139">O método `addItemAttachmentAsync` anexa o item com o identificador especificado do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="f6309-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f6309-805">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f6309-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f6309-806">Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a outros itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="f6309-806">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-807">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-807">Parameters:</span></span>

|<span data-ttu-id="f6309-808">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-808">Name</span></span>|<span data-ttu-id="f6309-809">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-809">Type</span></span>|<span data-ttu-id="f6309-810">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-810">Attributes</span></span>|<span data-ttu-id="f6309-811">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="f6309-812">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-812">String</span></span>||<span data-ttu-id="f6309-p140">O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f6309-815">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-815">String</span></span>||<span data-ttu-id="f6309-p141">O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f6309-818">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-818">Object</span></span>|<span data-ttu-id="f6309-819">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-819">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-820">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-821">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-821">Object</span></span>|<span data-ttu-id="f6309-822">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-822">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-823">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-824">function</span><span class="sxs-lookup"><span data-stu-id="f6309-824">function</span></span>|<span data-ttu-id="f6309-825">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-825">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-826">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f6309-827">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f6309-828">Se não for possível adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` com a descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f6309-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f6309-829">Erros</span><span class="sxs-lookup"><span data-stu-id="f6309-829">Errors</span></span>

|<span data-ttu-id="f6309-830">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f6309-830">Error code</span></span>|<span data-ttu-id="f6309-831">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f6309-832">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-833">Requirements</span></span>

|<span data-ttu-id="f6309-834">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-834">Requirement</span></span>|<span data-ttu-id="f6309-835">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-836">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-837">1.1</span><span class="sxs-lookup"><span data-stu-id="f6309-837">1.1</span></span>|
|[<span data-ttu-id="f6309-838">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-840">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-841">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-842">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-842">Example</span></span>

<span data-ttu-id="f6309-843">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f6309-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="f6309-844">close()</span><span class="sxs-lookup"><span data-stu-id="f6309-844">close()</span></span>

<span data-ttu-id="f6309-845">Fecha o item atual que está sendo redigido.</span><span class="sxs-lookup"><span data-stu-id="f6309-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f6309-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item possuir alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação de fechamento.</span><span class="sxs-lookup"><span data-stu-id="f6309-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-848">No Outlook na Web, se o item for um compromisso e tiver sido salvo anteriormente usando `saveAsync`, será solicitado ao usuário para salvar, descartar ou cancelar, mesmo que nenhuma alteração tenha ocorrido após o item ter sido salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="f6309-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f6309-849">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="f6309-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-850">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-850">Requirements</span></span>

|<span data-ttu-id="f6309-851">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-851">Requirement</span></span>|<span data-ttu-id="f6309-852">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-853">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-854">1.3</span><span class="sxs-lookup"><span data-stu-id="f6309-854">1.3</span></span>|
|[<span data-ttu-id="f6309-855">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-856">Restrito</span><span class="sxs-lookup"><span data-stu-id="f6309-856">Restricted</span></span>|
|[<span data-ttu-id="f6309-857">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-858">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f6309-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f6309-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f6309-860">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f6309-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-861">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-861">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-862">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f6309-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f6309-863">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f6309-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f6309-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f6309-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-867">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-867">Parameters:</span></span>

|<span data-ttu-id="f6309-868">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-868">Name</span></span>|<span data-ttu-id="f6309-869">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-869">Type</span></span>|<span data-ttu-id="f6309-870">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-870">Attributes</span></span>|<span data-ttu-id="f6309-871">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f6309-872">String | Object</span><span class="sxs-lookup"><span data-stu-id="f6309-872">String &#124; Object</span></span>||<span data-ttu-id="f6309-p144">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f6309-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f6309-875">**OU**</span><span class="sxs-lookup"><span data-stu-id="f6309-875">**OR**</span></span><br/><span data-ttu-id="f6309-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="f6309-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f6309-878">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-878">String</span></span>|<span data-ttu-id="f6309-879">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-879">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-p146">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f6309-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f6309-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f6309-883">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-883">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-884">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f6309-885">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-885">String</span></span>||<span data-ttu-id="f6309-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f6309-888">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-888">String</span></span>||<span data-ttu-id="f6309-889">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f6309-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f6309-890">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-890">String</span></span>||<span data-ttu-id="f6309-p148">Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f6309-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f6309-893">Booleano</span><span class="sxs-lookup"><span data-stu-id="f6309-893">Boolean</span></span>||<span data-ttu-id="f6309-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f6309-896">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-896">String</span></span>||<span data-ttu-id="f6309-p150">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f6309-900">função</span><span class="sxs-lookup"><span data-stu-id="f6309-900">function</span></span>|<span data-ttu-id="f6309-901">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-901">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-902">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-903">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-903">Requirements</span></span>

|<span data-ttu-id="f6309-904">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-904">Requirement</span></span>|<span data-ttu-id="f6309-905">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-906">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-907">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-907">1.0</span></span>|
|[<span data-ttu-id="f6309-908">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-909">ReadItem</span></span>|
|[<span data-ttu-id="f6309-910">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-911">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f6309-912">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f6309-912">Examples</span></span>

<span data-ttu-id="f6309-913">O código a seguir passa uma sequência de caracteres para a função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f6309-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f6309-914">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f6309-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f6309-915">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f6309-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f6309-916">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f6309-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f6309-917">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f6309-918">Resposta com o corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f6309-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f6309-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="f6309-920">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f6309-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-921">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-922">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f6309-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f6309-923">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f6309-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f6309-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f6309-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-927">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-927">Parameters:</span></span>

|<span data-ttu-id="f6309-928">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-928">Name</span></span>|<span data-ttu-id="f6309-929">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-929">Type</span></span>|<span data-ttu-id="f6309-930">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-930">Attributes</span></span>|<span data-ttu-id="f6309-931">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f6309-932">String | Object</span><span class="sxs-lookup"><span data-stu-id="f6309-932">String &#124; Object</span></span>||<span data-ttu-id="f6309-p152">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f6309-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f6309-935">**OU**</span><span class="sxs-lookup"><span data-stu-id="f6309-935">**OR**</span></span><br/><span data-ttu-id="f6309-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="f6309-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f6309-938">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-938">String</span></span>|<span data-ttu-id="f6309-939">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-939">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-p154">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f6309-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f6309-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f6309-943">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-943">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-944">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f6309-945">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-945">String</span></span>||<span data-ttu-id="f6309-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f6309-948">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-948">String</span></span>||<span data-ttu-id="f6309-949">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f6309-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f6309-950">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-950">String</span></span>||<span data-ttu-id="f6309-p156">Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f6309-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f6309-953">Booleano</span><span class="sxs-lookup"><span data-stu-id="f6309-953">Boolean</span></span>||<span data-ttu-id="f6309-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="f6309-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f6309-956">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-956">String</span></span>||<span data-ttu-id="f6309-p158">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f6309-960">função</span><span class="sxs-lookup"><span data-stu-id="f6309-960">function</span></span>|<span data-ttu-id="f6309-961">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-961">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-962">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-963">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-963">Requirements</span></span>

|<span data-ttu-id="f6309-964">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-964">Requirement</span></span>|<span data-ttu-id="f6309-965">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-966">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-967">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-967">1.0</span></span>|
|[<span data-ttu-id="f6309-968">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-969">ReadItem</span></span>|
|[<span data-ttu-id="f6309-970">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-971">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f6309-972">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f6309-972">Examples</span></span>

<span data-ttu-id="f6309-973">O código a seguir passa uma sequência de caracteres para a função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f6309-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f6309-974">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f6309-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f6309-975">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f6309-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f6309-976">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f6309-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f6309-977">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f6309-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f6309-978">Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f6309-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f6309-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f6309-980">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f6309-980">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-981">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-981">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-982">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-982">Requirements</span></span>

|<span data-ttu-id="f6309-983">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-983">Requirement</span></span>|<span data-ttu-id="f6309-984">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-985">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-986">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-986">1.0</span></span>|
|[<span data-ttu-id="f6309-987">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-988">ReadItem</span></span>|
|[<span data-ttu-id="f6309-989">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-990">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-991">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-991">Returns:</span></span>

<span data-ttu-id="f6309-992">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f6309-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f6309-993">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-993">Example</span></span>

<span data-ttu-id="f6309-994">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-994">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f6309-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f6309-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f6309-996">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f6309-996">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-997">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-997">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-998">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-998">Parameters:</span></span>

|<span data-ttu-id="f6309-999">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-999">Name</span></span>|<span data-ttu-id="f6309-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1000">Type</span></span>|<span data-ttu-id="f6309-1001">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="f6309-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f6309-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="f6309-1003">Um dos valores da enumeração EntityType.</span><span class="sxs-lookup"><span data-stu-id="f6309-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1004">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1004">Requirements</span></span>

|<span data-ttu-id="f6309-1005">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1005">Requirement</span></span>|<span data-ttu-id="f6309-1006">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1007">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-1008">1.0</span></span>|
|[<span data-ttu-id="f6309-1009">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1010">Restrito</span><span class="sxs-lookup"><span data-stu-id="f6309-1010">Restricted</span></span>|
|[<span data-ttu-id="f6309-1011">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1012">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1013">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1013">Returns:</span></span>

<span data-ttu-id="f6309-1014">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="f6309-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f6309-1015">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f6309-1015">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="f6309-1016">Caso contrário, o tipo dos objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f6309-1017">Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="f6309-1018">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f6309-1018">Value of `entityType`</span></span>|<span data-ttu-id="f6309-1019">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="f6309-1019">Type of objects in returned array</span></span>|<span data-ttu-id="f6309-1020">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="f6309-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="f6309-1021">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1021">String</span></span>|<span data-ttu-id="f6309-1022">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f6309-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="f6309-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="f6309-1023">Contact</span></span>|<span data-ttu-id="f6309-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f6309-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="f6309-1025">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1025">String</span></span>|<span data-ttu-id="f6309-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f6309-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="f6309-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f6309-1027">MeetingSuggestion</span></span>|<span data-ttu-id="f6309-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f6309-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="f6309-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f6309-1029">PhoneNumber</span></span>|<span data-ttu-id="f6309-1030">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f6309-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="f6309-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f6309-1031">TaskSuggestion</span></span>|<span data-ttu-id="f6309-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f6309-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="f6309-1033">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1033">String</span></span>|<span data-ttu-id="f6309-1034">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f6309-1034">**Restricted**</span></span>|

<span data-ttu-id="f6309-1035">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f6309-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f6309-1036">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1036">Example</span></span>

<span data-ttu-id="f6309-1037">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa os endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f6309-1037">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f6309-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f6309-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f6309-1039">Retorna entidades conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f6309-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1040">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-1041">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor especificado no elemento `FilterName` .</span><span class="sxs-lookup"><span data-stu-id="f6309-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1042">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1042">Parameters:</span></span>

|<span data-ttu-id="f6309-1043">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1043">Name</span></span>|<span data-ttu-id="f6309-1044">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1044">Type</span></span>|<span data-ttu-id="f6309-1045">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f6309-1046">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1046">String</span></span>|<span data-ttu-id="f6309-1047">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="f6309-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1048">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1048">Requirements</span></span>

|<span data-ttu-id="f6309-1049">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1049">Requirement</span></span>|<span data-ttu-id="f6309-1050">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1051">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-1052">1.0</span></span>|
|[<span data-ttu-id="f6309-1053">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1054">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1055">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1056">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1057">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1057">Returns:</span></span>

<span data-ttu-id="f6309-p160">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f6309-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f6309-1060">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f6309-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="f6309-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="f6309-1062">Obtém dados de inicialização que são passados quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="f6309-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1063">Esse método só é compatível com o Outlook 2016 ou posterior para o Windows (versões Click-to-Run posteriores à 16.0.8413.1000) e o Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="f6309-1063">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1064">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1064">Parameters:</span></span>
|<span data-ttu-id="f6309-1065">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1065">Name</span></span>|<span data-ttu-id="f6309-1066">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1066">Type</span></span>|<span data-ttu-id="f6309-1067">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1067">Attributes</span></span>|<span data-ttu-id="f6309-1068">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f6309-1069">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1069">Object</span></span>|<span data-ttu-id="f6309-1070">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1071">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1072">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1072">Object</span></span>|<span data-ttu-id="f6309-1073">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1074">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-1075">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1075">function</span></span>|<span data-ttu-id="f6309-1076">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1077">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f6309-1078">Em caso de êxito, os dados de inicialização são fornecidos na propriedade `asyncResult.value` como uma sequência de caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-1078">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="f6309-1079">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1080">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1080">Requirements</span></span>

|<span data-ttu-id="f6309-1081">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1081">Requirement</span></span>|<span data-ttu-id="f6309-1082">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1083">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1084">Visualização</span><span class="sxs-lookup"><span data-stu-id="f6309-1084">Preview</span></span>|
|[<span data-ttu-id="f6309-1085">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1086">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1087">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1088">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-1089">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1089">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="f6309-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f6309-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f6309-1091">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f6309-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1092">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-1092">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-p161">O método `getRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="f6309-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f6309-1096">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="f6309-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f6309-1097">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f6309-p162">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="f6309-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-1101">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1101">Requirements</span></span>

|<span data-ttu-id="f6309-1102">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1102">Requirement</span></span>|<span data-ttu-id="f6309-1103">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1104">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1104">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-1105">1.0</span></span>|
|[<span data-ttu-id="f6309-1106">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1107">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1108">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1109">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1110">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1110">Returns:</span></span>

<span data-ttu-id="f6309-p163">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="f6309-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f6309-1113">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f6309-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f6309-1114">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f6309-1115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1115">Example</span></span>

<span data-ttu-id="f6309-1116">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="f6309-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f6309-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f6309-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f6309-1118">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f6309-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1119">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-1119">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-1120">O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="f6309-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f6309-p164">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="f6309-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1123">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1123">Parameters:</span></span>

|<span data-ttu-id="f6309-1124">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1124">Name</span></span>|<span data-ttu-id="f6309-1125">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1125">Type</span></span>|<span data-ttu-id="f6309-1126">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f6309-1127">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1127">String</span></span>|<span data-ttu-id="f6309-1128">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="f6309-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1129">Requirements</span></span>

|<span data-ttu-id="f6309-1130">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1130">Requirement</span></span>|<span data-ttu-id="f6309-1131">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1132">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-1133">1.0</span></span>|
|[<span data-ttu-id="f6309-1134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1135">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1136">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1137">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1138">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1138">Returns:</span></span>

<span data-ttu-id="f6309-1139">Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f6309-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f6309-1140">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f6309-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f6309-1141">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f6309-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f6309-1142">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f6309-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f6309-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f6309-1144">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f6309-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f6309-p165">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f6309-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1147">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1147">Parameters:</span></span>

|<span data-ttu-id="f6309-1148">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1148">Name</span></span>|<span data-ttu-id="f6309-1149">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1149">Type</span></span>|<span data-ttu-id="f6309-1150">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1150">Attributes</span></span>|<span data-ttu-id="f6309-1151">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="f6309-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f6309-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f6309-p166">Solicita um formato para os dados. Se for Text, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="f6309-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="f6309-1156">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1156">Object</span></span>|<span data-ttu-id="f6309-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1158">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1159">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1159">Object</span></span>|<span data-ttu-id="f6309-1160">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1161">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-1162">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1162">function</span></span>||<span data-ttu-id="f6309-1163">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f6309-1164">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f6309-1165">Para acessar a propriedade de origem de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1165">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1166">Requirements</span></span>

|<span data-ttu-id="f6309-1167">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1167">Requirement</span></span>|<span data-ttu-id="f6309-1168">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1169">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="f6309-1170">1.2</span></span>|
|[<span data-ttu-id="f6309-1171">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-1173">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1174">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1175">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1175">Returns:</span></span>

<span data-ttu-id="f6309-1176">Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f6309-1177">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f6309-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f6309-1178">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f6309-1179">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f6309-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f6309-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f6309-p168">Obtém as entidades encontradas em uma correspondência destacada que um usuário selecionou. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f6309-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1183">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-1183">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-1184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1184">Requirements</span></span>

|<span data-ttu-id="f6309-1185">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1185">Requirement</span></span>|<span data-ttu-id="f6309-1186">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1187">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="f6309-1188">-16</span></span>|
|[<span data-ttu-id="f6309-1189">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1190">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1191">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1192">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1193">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1193">Returns:</span></span>

<span data-ttu-id="f6309-1194">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f6309-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f6309-1195">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1195">Example</span></span>

<span data-ttu-id="f6309-1196">O exemplo a seguir acessa as entidades de endereços na correspondência destacada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="f6309-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f6309-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f6309-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f6309-p169">Retorna valores do tipo sequência de caracteres em uma correspondência destacada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências destacadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="f6309-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1200">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f6309-1200">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f6309-p170">O método `getSelectedRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="f6309-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f6309-1204">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="f6309-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f6309-1205">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f6309-p171">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar o corpo e não tentar retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="f6309-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6309-1209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1209">Requirements</span></span>

|<span data-ttu-id="f6309-1210">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1210">Requirement</span></span>|<span data-ttu-id="f6309-1211">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1212">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="f6309-1213">-16</span></span>|
|[<span data-ttu-id="f6309-1214">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1215">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1216">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1217">Leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f6309-1218">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f6309-1218">Returns:</span></span>

<span data-ttu-id="f6309-p172">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="f6309-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f6309-1221">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1221">Example</span></span>

<span data-ttu-id="f6309-1222">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="f6309-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="f6309-1223">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f6309-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="f6309-1224">Obtém as propriedades do compromisso ou da mensagem selecionada em uma pasta compartilhada, um calendário ou uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="f6309-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1225">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1225">Parameters:</span></span>

|<span data-ttu-id="f6309-1226">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1226">Name</span></span>|<span data-ttu-id="f6309-1227">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1227">Type</span></span>|<span data-ttu-id="f6309-1228">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1228">Attributes</span></span>|<span data-ttu-id="f6309-1229">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f6309-1230">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1230">Object</span></span>|<span data-ttu-id="f6309-1231">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1232">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1233">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1233">Object</span></span>|<span data-ttu-id="f6309-1234">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1235">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-1236">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1236">function</span></span>||<span data-ttu-id="f6309-1237">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f6309-1238">As propriedades compartilhadas são fornecidas como um objeto [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1238">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f6309-1239">Esse objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="f6309-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1240">Requirements</span></span>

|<span data-ttu-id="f6309-1241">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1241">Requirement</span></span>|<span data-ttu-id="f6309-1242">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1243">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1244">Visualização</span><span class="sxs-lookup"><span data-stu-id="f6309-1244">Preview</span></span>|
|[<span data-ttu-id="f6309-1245">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1246">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1247">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1248">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-1249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f6309-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f6309-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f6309-1251">Carrega de forma assíncrona as propriedades personalizadas desse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f6309-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f6309-p174">As propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item e o suplemento atuais. As propriedades personalizadas não são criptografadas no item, portanto, isto não deve ser usado como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="f6309-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1255">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1255">Parameters:</span></span>

|<span data-ttu-id="f6309-1256">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1256">Name</span></span>|<span data-ttu-id="f6309-1257">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1257">Type</span></span>|<span data-ttu-id="f6309-1258">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1258">Attributes</span></span>|<span data-ttu-id="f6309-1259">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="f6309-1260">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1260">function</span></span>||<span data-ttu-id="f6309-1261">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f6309-1262">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f6309-1263">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas no servidor.</span><span class="sxs-lookup"><span data-stu-id="f6309-1263">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="f6309-1264">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1264">Object</span></span>|<span data-ttu-id="f6309-1265">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1266">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1266">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="f6309-1267">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1268">Requirements</span></span>

|<span data-ttu-id="f6309-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1269">Requirement</span></span>|<span data-ttu-id="f6309-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1271">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="f6309-1272">1.0</span></span>|
|[<span data-ttu-id="f6309-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1274">ReadItem</span></span>|
|[<span data-ttu-id="f6309-1275">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1276">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-1277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1277">Example</span></span>

<span data-ttu-id="f6309-p177">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f6309-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f6309-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f6309-1282">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f6309-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f6309-p178">O método `removeAttachmentAsync` remove do item o anexo com o identificador especificado. Conforme as práticas recomendadas, você deve usar o identificador do anexo para remover o anexo apenas se o mesmo aplicativo de email tiver inserido o anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador de anexos é válido somente dentro da mesma sessão. Uma sessão é considerada encerrada quando o usuário fecha o aplicativo, ou se o usuário começa a escrever um email em um formulário embutido e, em seguida, abre o mesmo formulário em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="f6309-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1287">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1287">Parameters:</span></span>

|<span data-ttu-id="f6309-1288">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1288">Name</span></span>|<span data-ttu-id="f6309-1289">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1289">Type</span></span>|<span data-ttu-id="f6309-1290">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1290">Attributes</span></span>|<span data-ttu-id="f6309-1291">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f6309-1292">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1292">String</span></span>||<span data-ttu-id="f6309-p179">O identificador do anexo a ser removido. O comprimento máximo da sequência de caracteres é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f6309-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="f6309-1295">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1295">Object</span></span>|<span data-ttu-id="f6309-1296">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1297">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1298">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1298">Object</span></span>|<span data-ttu-id="f6309-1299">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1300">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-1301">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1301">function</span></span>|<span data-ttu-id="f6309-1302">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1303">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f6309-1304">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="f6309-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f6309-1305">Erros</span><span class="sxs-lookup"><span data-stu-id="f6309-1305">Errors</span></span>

|<span data-ttu-id="f6309-1306">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f6309-1306">Error code</span></span>|<span data-ttu-id="f6309-1307">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="f6309-1308">O identificador do anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="f6309-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1309">Requirements</span></span>

|<span data-ttu-id="f6309-1310">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1310">Requirement</span></span>|<span data-ttu-id="f6309-1311">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1312">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="f6309-1313">1.1</span></span>|
|[<span data-ttu-id="f6309-1314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-1316">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1317">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-1318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1318">Example</span></span>

<span data-ttu-id="f6309-1319">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="f6309-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f6309-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f6309-1320">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f6309-1321">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="f6309-1321">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f6309-1322">Atualmente, os tipos de evento compatíveis são `Office.EventType.AppointmentTimeChanged` , `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="f6309-1322">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1323">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1323">Parameters:</span></span>

| <span data-ttu-id="f6309-1324">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1324">Name</span></span> | <span data-ttu-id="f6309-1325">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1325">Type</span></span> | <span data-ttu-id="f6309-1326">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1326">Attributes</span></span> | <span data-ttu-id="f6309-1327">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f6309-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f6309-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f6309-1329">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="f6309-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f6309-1330">Função</span><span class="sxs-lookup"><span data-stu-id="f6309-1330">Function</span></span> || <span data-ttu-id="f6309-p180">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="f6309-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f6309-1334">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1334">Object</span></span> | <span data-ttu-id="f6309-1335">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="f6309-1336">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f6309-1337">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1337">Object</span></span> | <span data-ttu-id="f6309-1338">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="f6309-1339">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f6309-1340">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1340">function</span></span>| <span data-ttu-id="f6309-1341">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1342">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1343">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1343">Requirements</span></span>

|<span data-ttu-id="f6309-1344">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1344">Requirement</span></span>| <span data-ttu-id="f6309-1345">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1346">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6309-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="f6309-1347">-17</span></span> |
|[<span data-ttu-id="f6309-1348">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6309-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1349">ReadItem</span></span> |
|[<span data-ttu-id="f6309-1350">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f6309-1351">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f6309-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f6309-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f6309-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="f6309-1353">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f6309-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="f6309-p181">Quando chamado, este método salva a mensagem atual como um rascunho e retorna o identificador do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook em modo de cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="f6309-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1357">Se o seu suplemento chamar `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver em modo de cache, poderá levar algum tempo antes do item realmente ser sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="f6309-1357">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="f6309-1358">Até que o item seja sincronizado, o uso de `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="f6309-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f6309-p183">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo de redação, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="f6309-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f6309-1362">Os seguintes clientes possuem um comportamento diferente para `saveAsync` em compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="f6309-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f6309-1363">O Outlook para Mac não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f6309-1363">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="f6309-1364">Chamar `saveAsync` em uma reunião no Outlook para Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="f6309-1364">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f6309-1365">O Outlook na Web sempre enviará um convite ou atualização quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f6309-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1366">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1366">Parameters:</span></span>

|<span data-ttu-id="f6309-1367">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1367">Name</span></span>|<span data-ttu-id="f6309-1368">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1368">Type</span></span>|<span data-ttu-id="f6309-1369">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1369">Attributes</span></span>|<span data-ttu-id="f6309-1370">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f6309-1371">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1371">Object</span></span>|<span data-ttu-id="f6309-1372">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1373">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1374">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1374">Object</span></span>|<span data-ttu-id="f6309-1375">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1376">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f6309-1377">function</span><span class="sxs-lookup"><span data-stu-id="f6309-1377">function</span></span>||<span data-ttu-id="f6309-1378">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f6309-1379">Em caso de sucesso, o identificador do item será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f6309-1379">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1380">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1380">Requirements</span></span>

|<span data-ttu-id="f6309-1381">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1381">Requirement</span></span>|<span data-ttu-id="f6309-1382">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1383">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="f6309-1384">1.3</span></span>|
|[<span data-ttu-id="f6309-1385">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-1387">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1388">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f6309-1389">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f6309-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f6309-p185">A seguir apresentamos um exemplo do parâmetro `result` passado para a função de retorno de chamada. A propriedade `value` contém o ID do item.</span><span class="sxs-lookup"><span data-stu-id="f6309-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f6309-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f6309-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f6309-1393">Insere dados no corpo ou no assunto de uma mensagem de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f6309-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f6309-p186">O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou no assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou do assunto, um erro será retornado. Após a inserção, o cursor será posicionado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="f6309-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f6309-1397">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f6309-1397">Parameters:</span></span>

|<span data-ttu-id="f6309-1398">Nome</span><span class="sxs-lookup"><span data-stu-id="f6309-1398">Name</span></span>|<span data-ttu-id="f6309-1399">Tipo</span><span class="sxs-lookup"><span data-stu-id="f6309-1399">Type</span></span>|<span data-ttu-id="f6309-1400">Atributos</span><span class="sxs-lookup"><span data-stu-id="f6309-1400">Attributes</span></span>|<span data-ttu-id="f6309-1401">Descrição</span><span class="sxs-lookup"><span data-stu-id="f6309-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="f6309-1402">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f6309-1402">String</span></span>||<span data-ttu-id="f6309-p187">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="f6309-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="f6309-1406">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1406">Object</span></span>|<span data-ttu-id="f6309-1407">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1408">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f6309-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f6309-1409">Objeto</span><span class="sxs-lookup"><span data-stu-id="f6309-1409">Object</span></span>|<span data-ttu-id="f6309-1410">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-1411">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f6309-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f6309-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f6309-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f6309-1413">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f6309-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="f6309-p188">Se for `text` , o estilo atual será aplicado no Outlook Web App e no Outlook. Se o campo for um editor HTML, somente os dados de texto serão inseridos, mesmo que os dados estejam em HTML.</span><span class="sxs-lookup"><span data-stu-id="f6309-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f6309-p189">Se for `html` e o campo for compatível com HTML (e o assunto não), o estilo atual será aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, um erro `InvalidDataFormat` será retornado.</span><span class="sxs-lookup"><span data-stu-id="f6309-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f6309-1418">Se `coercionType` não estiver definido, o resultado dependerá do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="f6309-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="f6309-1419">função</span><span class="sxs-lookup"><span data-stu-id="f6309-1419">function</span></span>||<span data-ttu-id="f6309-1420">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f6309-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f6309-1421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f6309-1421">Requirements</span></span>

|<span data-ttu-id="f6309-1422">Requisito</span><span class="sxs-lookup"><span data-stu-id="f6309-1422">Requirement</span></span>|<span data-ttu-id="f6309-1423">Valor</span><span class="sxs-lookup"><span data-stu-id="f6309-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6309-1424">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f6309-1424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f6309-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="f6309-1425">1.2</span></span>|
|[<span data-ttu-id="f6309-1426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f6309-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f6309-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f6309-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="f6309-1428">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f6309-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="f6309-1429">Redação</span><span class="sxs-lookup"><span data-stu-id="f6309-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f6309-1430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f6309-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```