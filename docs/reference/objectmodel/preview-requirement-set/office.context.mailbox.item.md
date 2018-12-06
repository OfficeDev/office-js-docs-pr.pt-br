
# <a name="item"></a><span data-ttu-id="a6266-101">item</span><span class="sxs-lookup"><span data-stu-id="a6266-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a6266-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a6266-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a6266-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a6266-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-105">Requirements</span></span>

|<span data-ttu-id="a6266-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-106">Requirement</span></span>|<span data-ttu-id="a6266-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-109">1.0</span></span>|
|[<span data-ttu-id="a6266-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="a6266-111">Restricted</span></span>|
|[<span data-ttu-id="a6266-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a6266-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a6266-114">Members and methods</span></span>

| <span data-ttu-id="a6266-115">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-115">Member</span></span> | <span data-ttu-id="a6266-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a6266-117">attachments</span><span class="sxs-lookup"><span data-stu-id="a6266-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a6266-118">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-118">Member</span></span> |
| [<span data-ttu-id="a6266-119">bcc</span><span class="sxs-lookup"><span data-stu-id="a6266-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a6266-120">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-120">Member</span></span> |
| [<span data-ttu-id="a6266-121">body</span><span class="sxs-lookup"><span data-stu-id="a6266-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="a6266-122">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-122">Member</span></span> |
| [<span data-ttu-id="a6266-123">cc</span><span class="sxs-lookup"><span data-stu-id="a6266-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a6266-124">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-124">Member</span></span> |
| [<span data-ttu-id="a6266-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="a6266-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a6266-126">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-126">Member</span></span> |
| [<span data-ttu-id="a6266-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a6266-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a6266-128">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-128">Member</span></span> |
| [<span data-ttu-id="a6266-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a6266-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a6266-130">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-130">Member</span></span> |
| [<span data-ttu-id="a6266-131">end</span><span class="sxs-lookup"><span data-stu-id="a6266-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a6266-132">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-132">Member</span></span> |
| [<span data-ttu-id="a6266-133">from</span><span class="sxs-lookup"><span data-stu-id="a6266-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="a6266-134">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-134">Member</span></span> |
| [<span data-ttu-id="a6266-135">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="a6266-135">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="a6266-136">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-136">Member</span></span> |
| [<span data-ttu-id="a6266-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a6266-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a6266-138">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-138">Member</span></span> |
| [<span data-ttu-id="a6266-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="a6266-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a6266-140">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-140">Member</span></span> |
| [<span data-ttu-id="a6266-141">itemId</span><span class="sxs-lookup"><span data-stu-id="a6266-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a6266-142">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-142">Member</span></span> |
| [<span data-ttu-id="a6266-143">itemType</span><span class="sxs-lookup"><span data-stu-id="a6266-143">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="a6266-144">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-144">Member</span></span> |
| [<span data-ttu-id="a6266-145">location</span><span class="sxs-lookup"><span data-stu-id="a6266-145">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="a6266-146">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-146">Member</span></span> |
| [<span data-ttu-id="a6266-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a6266-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a6266-148">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-148">Member</span></span> |
| [<span data-ttu-id="a6266-149">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a6266-149">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="a6266-150">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-150">Member</span></span> |
| [<span data-ttu-id="a6266-151">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a6266-151">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a6266-152">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-152">Member</span></span> |
| [<span data-ttu-id="a6266-153">organizer</span><span class="sxs-lookup"><span data-stu-id="a6266-153">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="a6266-154">Member</span><span class="sxs-lookup"><span data-stu-id="a6266-154">Member</span></span> |
| [<span data-ttu-id="a6266-155">recurrence</span><span class="sxs-lookup"><span data-stu-id="a6266-155">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="a6266-156">Member</span><span class="sxs-lookup"><span data-stu-id="a6266-156">Member</span></span> |
| [<span data-ttu-id="a6266-157">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a6266-157">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a6266-158">Member</span><span class="sxs-lookup"><span data-stu-id="a6266-158">Member</span></span> |
| [<span data-ttu-id="a6266-159">sender</span><span class="sxs-lookup"><span data-stu-id="a6266-159">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="a6266-160">Member</span><span class="sxs-lookup"><span data-stu-id="a6266-160">Member</span></span> |
| [<span data-ttu-id="a6266-161">seriesId</span><span class="sxs-lookup"><span data-stu-id="a6266-161">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a6266-162">Member</span><span class="sxs-lookup"><span data-stu-id="a6266-162">Member</span></span> |
| [<span data-ttu-id="a6266-163">start</span><span class="sxs-lookup"><span data-stu-id="a6266-163">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a6266-164">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-164">Member</span></span> |
| [<span data-ttu-id="a6266-165">subject</span><span class="sxs-lookup"><span data-stu-id="a6266-165">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="a6266-166">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-166">Member</span></span> |
| [<span data-ttu-id="a6266-167">to</span><span class="sxs-lookup"><span data-stu-id="a6266-167">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a6266-168">Membro</span><span class="sxs-lookup"><span data-stu-id="a6266-168">Member</span></span> |
| [<span data-ttu-id="a6266-169">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-169">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a6266-170">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-170">Method</span></span> |
| [<span data-ttu-id="a6266-171">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a6266-171">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a6266-172">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-172">Method</span></span> |
| [<span data-ttu-id="a6266-173">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-173">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a6266-174">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-174">Method</span></span> |
| [<span data-ttu-id="a6266-175">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-175">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a6266-176">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-176">Method</span></span> |
| [<span data-ttu-id="a6266-177">close</span><span class="sxs-lookup"><span data-stu-id="a6266-177">close</span></span>](#close) | <span data-ttu-id="a6266-178">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-178">Method</span></span> |
| [<span data-ttu-id="a6266-179">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a6266-179">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a6266-180">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-180">Method</span></span> |
| [<span data-ttu-id="a6266-181">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a6266-181">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a6266-182">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-182">Method</span></span> |
| [<span data-ttu-id="a6266-183">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-183">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="a6266-184">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-184">Method</span></span> |
| [<span data-ttu-id="a6266-185">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-185">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a6266-186">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-186">Method</span></span> |
| [<span data-ttu-id="a6266-187">getEntities</span><span class="sxs-lookup"><span data-stu-id="a6266-187">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a6266-188">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-188">Method</span></span> |
| [<span data-ttu-id="a6266-189">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a6266-189">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a6266-190">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-190">Method</span></span> |
| [<span data-ttu-id="a6266-191">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a6266-191">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a6266-192">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-192">Method</span></span> |
| [<span data-ttu-id="a6266-193">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-193">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a6266-194">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-194">Method</span></span> |
| [<span data-ttu-id="a6266-195">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a6266-195">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a6266-196">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-196">Method</span></span> |
| [<span data-ttu-id="a6266-197">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a6266-197">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a6266-198">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-198">Method</span></span> |
| [<span data-ttu-id="a6266-199">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-199">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a6266-200">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-200">Method</span></span> |
| [<span data-ttu-id="a6266-201">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a6266-201">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a6266-202">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-202">Method</span></span> |
| [<span data-ttu-id="a6266-203">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a6266-203">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a6266-204">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-204">Method</span></span> |
| [<span data-ttu-id="a6266-205">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-205">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="a6266-206">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-206">Method</span></span> |
| [<span data-ttu-id="a6266-207">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-207">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a6266-208">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-208">Method</span></span> |
| [<span data-ttu-id="a6266-209">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-209">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a6266-210">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-210">Method</span></span> |
| [<span data-ttu-id="a6266-211">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-211">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a6266-212">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-212">Method</span></span> |
| [<span data-ttu-id="a6266-213">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-213">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a6266-214">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-214">Method</span></span> |
| [<span data-ttu-id="a6266-215">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a6266-215">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a6266-216">Método</span><span class="sxs-lookup"><span data-stu-id="a6266-216">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a6266-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-217">Example</span></span>

<span data-ttu-id="a6266-218">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6266-218">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="a6266-219">Membros</span><span class="sxs-lookup"><span data-stu-id="a6266-219">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a6266-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6266-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a6266-221">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="a6266-221">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a6266-222">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-223">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="a6266-223">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a6266-224">Para saber mais, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a6266-224">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-225">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-225">Type:</span></span>

*   <span data-ttu-id="a6266-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6266-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-227">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-227">Requirements</span></span>

|<span data-ttu-id="a6266-228">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-228">Requirement</span></span>|<span data-ttu-id="a6266-229">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-230">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-231">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-231">1.0</span></span>|
|[<span data-ttu-id="a6266-232">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-232">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-233">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-233">ReadItem</span></span>|
|[<span data-ttu-id="a6266-234">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-234">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-235">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-235">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-236">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-236">Example</span></span>

<span data-ttu-id="a6266-237">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-237">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a6266-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a6266-239">Obtém um objeto que fornece métodos para acessar ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-239">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a6266-240">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a6266-240">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-241">Type:</span></span>

*   [<span data-ttu-id="a6266-242">Destinatários</span><span class="sxs-lookup"><span data-stu-id="a6266-242">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a6266-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-243">Requirements</span></span>

|<span data-ttu-id="a6266-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-244">Requirement</span></span>|<span data-ttu-id="a6266-245">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-246">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-247">1.1</span><span class="sxs-lookup"><span data-stu-id="a6266-247">1.1</span></span>|
|[<span data-ttu-id="a6266-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-249">ReadItem</span></span>|
|[<span data-ttu-id="a6266-250">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-251">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-251">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-252">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a6266-253">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a6266-253">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a6266-254">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="a6266-254">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-255">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-255">Type:</span></span>

*   [<span data-ttu-id="a6266-256">Corpo</span><span class="sxs-lookup"><span data-stu-id="a6266-256">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a6266-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-257">Requirements</span></span>

|<span data-ttu-id="a6266-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-258">Requirement</span></span>|<span data-ttu-id="a6266-259">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-261">1.1</span><span class="sxs-lookup"><span data-stu-id="a6266-261">1.1</span></span>|
|[<span data-ttu-id="a6266-262">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-263">ReadItem</span></span>|
|[<span data-ttu-id="a6266-264">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-265">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-265">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a6266-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a6266-267">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-267">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a6266-268">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-268">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-269">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-269">Read mode</span></span>

<span data-ttu-id="a6266-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a6266-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-272">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-272">Compose mode</span></span>

<span data-ttu-id="a6266-273">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-273">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-274">Type:</span></span>

*   <span data-ttu-id="a6266-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-276">Requirements</span></span>

|<span data-ttu-id="a6266-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-277">Requirement</span></span>|<span data-ttu-id="a6266-278">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-279">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-280">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-280">1.0</span></span>|
|[<span data-ttu-id="a6266-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-282">ReadItem</span></span>|
|[<span data-ttu-id="a6266-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-284">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-284">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-285">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a6266-286">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a6266-286">(nullable) conversationId :String</span></span>

<span data-ttu-id="a6266-287">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="a6266-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a6266-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="a6266-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a6266-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="a6266-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-292">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-292">Type:</span></span>

*   <span data-ttu-id="a6266-293">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-294">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-294">Requirements</span></span>

|<span data-ttu-id="a6266-295">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-295">Requirement</span></span>|<span data-ttu-id="a6266-296">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-297">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-298">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-298">1.0</span></span>|
|[<span data-ttu-id="a6266-299">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-299">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-300">ReadItem</span></span>|
|[<span data-ttu-id="a6266-301">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-301">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-302">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-302">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a6266-303">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a6266-303">dateTimeCreated :Date</span></span>

<span data-ttu-id="a6266-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-306">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-306">Type:</span></span>

*   <span data-ttu-id="a6266-307">Data</span><span class="sxs-lookup"><span data-stu-id="a6266-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-308">Requirements</span></span>

|<span data-ttu-id="a6266-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-309">Requirement</span></span>|<span data-ttu-id="a6266-310">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-312">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-312">1.0</span></span>|
|[<span data-ttu-id="a6266-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-314">ReadItem</span></span>|
|[<span data-ttu-id="a6266-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-316">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-317">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a6266-318">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a6266-318">dateTimeModified :Date</span></span>

<span data-ttu-id="a6266-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-321">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-321">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-322">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-322">Type:</span></span>

*   <span data-ttu-id="a6266-323">Data</span><span class="sxs-lookup"><span data-stu-id="a6266-323">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-324">Requirements</span></span>

|<span data-ttu-id="a6266-325">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-325">Requirement</span></span>|<span data-ttu-id="a6266-326">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-328">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-328">1.0</span></span>|
|[<span data-ttu-id="a6266-329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-330">ReadItem</span></span>|
|[<span data-ttu-id="a6266-331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-332">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-333">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a6266-334">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6266-334">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a6266-335">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="a6266-335">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a6266-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a6266-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-338">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-338">Read mode</span></span>

<span data-ttu-id="a6266-339">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a6266-339">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-340">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-340">Compose mode</span></span>

<span data-ttu-id="a6266-341">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a6266-341">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a6266-342">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a6266-342">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-343">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-343">Type:</span></span>

*   <span data-ttu-id="a6266-344">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6266-344">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-345">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-345">Requirements</span></span>

|<span data-ttu-id="a6266-346">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-346">Requirement</span></span>|<span data-ttu-id="a6266-347">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-348">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-349">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-349">1.0</span></span>|
|[<span data-ttu-id="a6266-350">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-351">ReadItem</span></span>|
|[<span data-ttu-id="a6266-352">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-353">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-354">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-354">Example</span></span>

<span data-ttu-id="a6266-355">O exemplo a seguir define a hora de término de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a6266-355">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a6266-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a6266-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a6266-357">Obtém o endereço de email do remetente de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a6266-p112">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a6266-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-360">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a6266-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-361">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-361">Read mode</span></span>

<span data-ttu-id="a6266-362">A propriedade `from` retorna um objeto `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="a6266-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="a6266-363">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-363">Compose mode</span></span>

<span data-ttu-id="a6266-364">A propriedade `from` retorna um objeto `From` que fornece um método para obtenção do valor de from.</span><span class="sxs-lookup"><span data-stu-id="a6266-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a6266-365">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-365">Type:</span></span>

*   <span data-ttu-id="a6266-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a6266-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-367">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-367">Requirements</span></span>

|<span data-ttu-id="a6266-368">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a6266-369">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-370">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-370">1.0</span></span>|<span data-ttu-id="a6266-371">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-371">1.7</span></span>|
|[<span data-ttu-id="a6266-372">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-373">ReadItem</span></span>|<span data-ttu-id="a6266-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-375">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-375">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-376">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-376">Read</span></span>|<span data-ttu-id="a6266-377">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-377">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="a6266-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="a6266-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="a6266-379">Obtém ou define os cabeçalhos de internet de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-379">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-380">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-380">Type:</span></span>

*   [<span data-ttu-id="a6266-381">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="a6266-381">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="a6266-382">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-382">Requirements</span></span>

|<span data-ttu-id="a6266-383">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-383">Requirement</span></span>|<span data-ttu-id="a6266-384">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-384">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-385">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-386">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-386">Preview</span></span>|
|[<span data-ttu-id="a6266-387">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-387">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-388">ReadItem</span></span>|
|[<span data-ttu-id="a6266-389">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-389">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-390">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-390">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a6266-391">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a6266-391">internetMessageId :String</span></span>

<span data-ttu-id="a6266-p113">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-394">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-394">Type:</span></span>

*   <span data-ttu-id="a6266-395">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-396">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-396">Requirements</span></span>

|<span data-ttu-id="a6266-397">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-397">Requirement</span></span>|<span data-ttu-id="a6266-398">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-399">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-400">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-400">1.0</span></span>|
|[<span data-ttu-id="a6266-401">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-402">ReadItem</span></span>|
|[<span data-ttu-id="a6266-403">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-404">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-405">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-405">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a6266-406">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a6266-406">itemClass :String</span></span>

<span data-ttu-id="a6266-p114">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a6266-p115">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a6266-411">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-411">Type</span></span>|<span data-ttu-id="a6266-412">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-412">Description</span></span>|<span data-ttu-id="a6266-413">classe de item</span><span class="sxs-lookup"><span data-stu-id="a6266-413">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a6266-414">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="a6266-414">Appointment items</span></span>|<span data-ttu-id="a6266-415">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a6266-415">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="a6266-416">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="a6266-416">Message items</span></span>|<span data-ttu-id="a6266-417">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="a6266-417">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a6266-418">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="a6266-418">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-419">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-419">Type:</span></span>

*   <span data-ttu-id="a6266-420">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-421">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-421">Requirements</span></span>

|<span data-ttu-id="a6266-422">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-422">Requirement</span></span>|<span data-ttu-id="a6266-423">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-424">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-425">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-425">1.0</span></span>|
|[<span data-ttu-id="a6266-426">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-427">ReadItem</span></span>|
|[<span data-ttu-id="a6266-428">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-429">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-430">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-430">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a6266-431">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a6266-431">(nullable) itemId :String</span></span>

<span data-ttu-id="a6266-p116">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-434">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a6266-434">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a6266-435">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6266-435">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a6266-436">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a6266-436">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a6266-437">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a6266-437">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a6266-p118">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-440">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-440">Type:</span></span>

*   <span data-ttu-id="a6266-441">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-441">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-442">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-442">Requirements</span></span>

|<span data-ttu-id="a6266-443">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-443">Requirement</span></span>|<span data-ttu-id="a6266-444">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-445">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-446">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-446">1.0</span></span>|
|[<span data-ttu-id="a6266-447">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-448">ReadItem</span></span>|
|[<span data-ttu-id="a6266-449">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-450">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-450">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-451">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-451">Example</span></span>

<span data-ttu-id="a6266-p119">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a6266-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a6266-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a6266-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a6266-455">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="a6266-455">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a6266-456">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-456">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-457">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-457">Type:</span></span>

*   [<span data-ttu-id="a6266-458">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a6266-458">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a6266-459">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-459">Requirements</span></span>

|<span data-ttu-id="a6266-460">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-460">Requirement</span></span>|<span data-ttu-id="a6266-461">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-462">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-463">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-463">1.0</span></span>|
|[<span data-ttu-id="a6266-464">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-465">ReadItem</span></span>|
|[<span data-ttu-id="a6266-466">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-467">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-467">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-468">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-468">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a6266-469">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a6266-469">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a6266-470">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-470">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-471">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-471">Read mode</span></span>

<span data-ttu-id="a6266-472">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-472">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-473">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-473">Compose mode</span></span>

<span data-ttu-id="a6266-474">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-474">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-475">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-475">Type:</span></span>

*   <span data-ttu-id="a6266-476">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a6266-476">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-477">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-477">Requirements</span></span>

|<span data-ttu-id="a6266-478">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-478">Requirement</span></span>|<span data-ttu-id="a6266-479">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-480">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-481">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-481">1.0</span></span>|
|[<span data-ttu-id="a6266-482">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-482">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-483">ReadItem</span></span>|
|[<span data-ttu-id="a6266-484">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-484">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-485">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-485">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-486">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-486">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a6266-487">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a6266-487">normalizedSubject :String</span></span>

<span data-ttu-id="a6266-p120">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a6266-p121">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="a6266-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-492">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-492">Type:</span></span>

*   <span data-ttu-id="a6266-493">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-493">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-494">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-494">Requirements</span></span>

|<span data-ttu-id="a6266-495">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-495">Requirement</span></span>|<span data-ttu-id="a6266-496">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-497">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-498">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-498">1.0</span></span>|
|[<span data-ttu-id="a6266-499">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-499">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-500">ReadItem</span></span>|
|[<span data-ttu-id="a6266-501">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-501">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-502">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-502">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-503">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-503">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a6266-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a6266-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a6266-505">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="a6266-505">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-506">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-506">Type:</span></span>

*   [<span data-ttu-id="a6266-507">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a6266-507">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a6266-508">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-508">Requirements</span></span>

|<span data-ttu-id="a6266-509">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-509">Requirement</span></span>|<span data-ttu-id="a6266-510">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-510">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-511">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-511">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-512">1.3</span><span class="sxs-lookup"><span data-stu-id="a6266-512">1.3</span></span>|
|[<span data-ttu-id="a6266-513">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-513">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-514">ReadItem</span></span>|
|[<span data-ttu-id="a6266-515">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-515">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-516">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-516">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a6266-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a6266-518">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="a6266-518">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a6266-519">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-519">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-520">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-520">Read mode</span></span>

<span data-ttu-id="a6266-521">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-521">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-522">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-522">Compose mode</span></span>

<span data-ttu-id="a6266-523">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-523">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-524">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-524">Type:</span></span>

*   <span data-ttu-id="a6266-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-526">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-526">Requirements</span></span>

|<span data-ttu-id="a6266-527">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-527">Requirement</span></span>|<span data-ttu-id="a6266-528">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-529">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-530">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-530">1.0</span></span>|
|[<span data-ttu-id="a6266-531">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-531">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-532">ReadItem</span></span>|
|[<span data-ttu-id="a6266-533">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-534">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-534">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-535">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-535">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a6266-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a6266-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a6266-537">Obtém o endereço de email do organizador para uma reunião especificada.</span><span class="sxs-lookup"><span data-stu-id="a6266-537">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-538">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-538">Read mode</span></span>

<span data-ttu-id="a6266-539">A propriedade `organizer` retorna um objeto [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) que representa o organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-539">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-540">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-540">Compose mode</span></span>

<span data-ttu-id="a6266-541">A propriedade `organizer` retorna um objeto [Organizer](/javascript/api/outlook/office.organizer) que fornece um método para obtenção do valor de organizer.</span><span class="sxs-lookup"><span data-stu-id="a6266-541">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-542">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-542">Type:</span></span>

*   <span data-ttu-id="a6266-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a6266-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-544">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-544">Requirements</span></span>

|<span data-ttu-id="a6266-545">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-545">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a6266-546">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-547">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-547">1.0</span></span>|<span data-ttu-id="a6266-548">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-548">1.7</span></span>|
|[<span data-ttu-id="a6266-549">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-549">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-550">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-550">ReadItem</span></span>|<span data-ttu-id="a6266-551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-551">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-552">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-553">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-553">Read</span></span>|<span data-ttu-id="a6266-554">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-555">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-555">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a6266-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a6266-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a6266-557">Obtém ou configura o padrão de recorrência de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-557">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a6266-558">Obtém o padrão de recorrência de uma solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-558">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a6266-559">Modos de leitura e redação para itens do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-559">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a6266-560">Modo de leitura para os itens de solicitação de reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-560">Read mode for meeting request items.</span></span>

<span data-ttu-id="a6266-561">A propriedade `recurrence` retorna um objeto [recurrence](/javascript/api/outlook/office.recurrence) para compromissos recorrentes ou solicitações de reuniões se um item for uma série ou uma instância em uma série.</span><span class="sxs-lookup"><span data-stu-id="a6266-561">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a6266-562">`null` retorna para compromissos individuais e solicitações de reunião de compromissos individuais.</span><span class="sxs-lookup"><span data-stu-id="a6266-562">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a6266-563">`undefined` retorna para mensagens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-563">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a6266-564">Observação: solicitações de reunião têm um valor `itemClass` de IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="a6266-564">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a6266-565">Observação: se o objeto de recorrência for `null`, isso indicará que o objeto é um compromisso individual ou uma solicitação de reunião de um compromisso individual e NÃO parte de uma série.</span><span class="sxs-lookup"><span data-stu-id="a6266-565">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-566">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-566">Type:</span></span>

* [<span data-ttu-id="a6266-567">Recurrence</span><span class="sxs-lookup"><span data-stu-id="a6266-567">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a6266-568">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-568">Requirement</span></span>|<span data-ttu-id="a6266-569">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-570">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-571">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-571">1.7</span></span>|
|[<span data-ttu-id="a6266-572">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-573">ReadItem</span></span>|
|[<span data-ttu-id="a6266-574">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-575">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-575">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a6266-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a6266-577">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="a6266-577">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a6266-578">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-578">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-579">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-579">Read mode</span></span>

<span data-ttu-id="a6266-580">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-580">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-581">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-581">Compose mode</span></span>

<span data-ttu-id="a6266-582">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-582">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-583">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-583">Type:</span></span>

*   <span data-ttu-id="a6266-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-585">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-585">Requirements</span></span>

|<span data-ttu-id="a6266-586">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-586">Requirement</span></span>|<span data-ttu-id="a6266-587">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-588">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-589">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-589">1.0</span></span>|
|[<span data-ttu-id="a6266-590">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-590">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-591">ReadItem</span></span>|
|[<span data-ttu-id="a6266-592">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-592">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-593">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-593">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-594">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-594">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a6266-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a6266-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a6266-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a6266-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a6266-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="a6266-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-600">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a6266-600">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-601">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-601">Type:</span></span>

*   [<span data-ttu-id="a6266-602">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a6266-602">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a6266-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-603">Requirements</span></span>

|<span data-ttu-id="a6266-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-604">Requirement</span></span>|<span data-ttu-id="a6266-605">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-607">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-607">1.0</span></span>|
|[<span data-ttu-id="a6266-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-609">ReadItem</span></span>|
|[<span data-ttu-id="a6266-610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-611">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-611">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-612">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-612">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a6266-613">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="a6266-613">(nullable) seriesId :String</span></span>

<span data-ttu-id="a6266-614">Obtém a id da série a qual uma instância pertence.</span><span class="sxs-lookup"><span data-stu-id="a6266-614">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a6266-615">No OWA e no Outlook, o `seriesId` retorna a ID dos Serviços Web do Exchange (EWS) do item pai (série) a qual este item pertence.</span><span class="sxs-lookup"><span data-stu-id="a6266-615">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a6266-616">No entanto, no iOS e no Android, o `seriesId` retorna a ID REST do item pai.</span><span class="sxs-lookup"><span data-stu-id="a6266-616">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-617">O identificador retornado pela propriedade `seriesId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a6266-617">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a6266-618">A propriedade `seriesId` não é idêntica à ID do Outlook usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6266-618">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a6266-619">Antes de fazer chamadas API REST usando esse valor, ela deverá ser convertida usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a6266-619">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a6266-620">Para obter detalhes, confira [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a6266-620">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a6266-621">A propriedade `seriesId` retorna `null` para itens que não têm itens pai como compromissos individuais, itens de série ou solicitações de reunião e retorna `undefined` para outros itens que não são solicitações de reunião.</span><span class="sxs-lookup"><span data-stu-id="a6266-621">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-622">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-622">Type:</span></span>

* <span data-ttu-id="a6266-623">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6266-623">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-624">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-624">Requirements</span></span>

|<span data-ttu-id="a6266-625">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-625">Requirement</span></span>|<span data-ttu-id="a6266-626">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-627">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-628">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-628">1.7</span></span>|
|[<span data-ttu-id="a6266-629">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-630">ReadItem</span></span>|
|[<span data-ttu-id="a6266-631">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-632">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-633">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-633">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a6266-634">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6266-634">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a6266-635">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="a6266-635">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a6266-p130">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a6266-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-638">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-638">Read mode</span></span>

<span data-ttu-id="a6266-639">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="a6266-639">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-640">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-640">Compose mode</span></span>

<span data-ttu-id="a6266-641">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a6266-641">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a6266-642">Ao usar o método [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="a6266-642">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-643">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-643">Type:</span></span>

*   <span data-ttu-id="a6266-644">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6266-644">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-645">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-645">Requirements</span></span>

|<span data-ttu-id="a6266-646">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-646">Requirement</span></span>|<span data-ttu-id="a6266-647">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-648">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-649">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-649">1.0</span></span>|
|[<span data-ttu-id="a6266-650">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-651">ReadItem</span></span>|
|[<span data-ttu-id="a6266-652">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-653">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-653">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-654">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-654">Example</span></span>

<span data-ttu-id="a6266-655">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="a6266-655">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a6266-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a6266-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a6266-657">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="a6266-657">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a6266-658">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="a6266-658">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-659">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-659">Read mode</span></span>

<span data-ttu-id="a6266-p131">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a6266-662">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-662">Compose mode</span></span>

<span data-ttu-id="a6266-663">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="a6266-663">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a6266-664">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-664">Type:</span></span>

*   <span data-ttu-id="a6266-665">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a6266-665">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-666">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-666">Requirements</span></span>

|<span data-ttu-id="a6266-667">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-667">Requirement</span></span>|<span data-ttu-id="a6266-668">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-669">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-670">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-670">1.0</span></span>|
|[<span data-ttu-id="a6266-671">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-672">ReadItem</span></span>|
|[<span data-ttu-id="a6266-673">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-674">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-674">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a6266-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a6266-676">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-676">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a6266-677">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-677">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6266-678">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-678">Read mode</span></span>

<span data-ttu-id="a6266-p133">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="a6266-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6266-681">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="a6266-681">Compose mode</span></span>

<span data-ttu-id="a6266-682">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-682">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a6266-683">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a6266-683">Type:</span></span>

*   <span data-ttu-id="a6266-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6266-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-685">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-685">Requirements</span></span>

|<span data-ttu-id="a6266-686">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-686">Requirement</span></span>|<span data-ttu-id="a6266-687">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-688">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-689">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-689">1.0</span></span>|
|[<span data-ttu-id="a6266-690">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-690">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-691">ReadItem</span></span>|
|[<span data-ttu-id="a6266-692">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-692">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-693">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-693">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-694">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-694">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a6266-695">Métodos</span><span class="sxs-lookup"><span data-stu-id="a6266-695">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a6266-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a6266-697">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="a6266-697">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a6266-698">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a6266-698">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a6266-699">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-699">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-700">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-700">Parameters:</span></span>
|<span data-ttu-id="a6266-701">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-701">Name</span></span>|<span data-ttu-id="a6266-702">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-702">Type</span></span>|<span data-ttu-id="a6266-703">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-703">Attributes</span></span>|<span data-ttu-id="a6266-704">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-704">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a6266-705">String</span><span class="sxs-lookup"><span data-stu-id="a6266-705">String</span></span>||<span data-ttu-id="a6266-p134">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a6266-708">String</span><span class="sxs-lookup"><span data-stu-id="a6266-708">String</span></span>||<span data-ttu-id="a6266-p135">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a6266-711">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-711">Object</span></span>|<span data-ttu-id="a6266-712">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-712">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-713">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-713">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-714">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-714">Object</span></span>|<span data-ttu-id="a6266-715">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-715">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-716">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-716">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a6266-717">Booliano</span><span class="sxs-lookup"><span data-stu-id="a6266-717">Boolean</span></span>|<span data-ttu-id="a6266-718">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-718">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-719">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-719">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a6266-720">function</span><span class="sxs-lookup"><span data-stu-id="a6266-720">function</span></span>|<span data-ttu-id="a6266-721">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-721">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-722">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-722">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6266-723">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-723">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a6266-724">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a6266-724">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6266-725">Erros</span><span class="sxs-lookup"><span data-stu-id="a6266-725">Errors</span></span>

|<span data-ttu-id="a6266-726">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a6266-726">Error code</span></span>|<span data-ttu-id="a6266-727">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-727">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a6266-728">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a6266-728">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a6266-729">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a6266-729">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a6266-730">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-730">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-731">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-731">Requirements</span></span>

|<span data-ttu-id="a6266-732">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-732">Requirement</span></span>|<span data-ttu-id="a6266-733">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-734">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-735">1.1</span><span class="sxs-lookup"><span data-stu-id="a6266-735">1.1</span></span>|
|[<span data-ttu-id="a6266-736">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-736">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-737">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-737">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-738">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-738">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-739">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-739">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6266-740">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a6266-740">Examples</span></span>

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

<span data-ttu-id="a6266-741">O exemplo a seguir adiciona um arquivo de imagem como um anexo embutido e faz referência ao anexo no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-741">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a6266-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a6266-743">Adiciona um arquivo a partir da codificação base64 a uma mensagem ou compromisso como anexo.</span><span class="sxs-lookup"><span data-stu-id="a6266-743">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a6266-744">O método `addFileAttachmentFromBase64Async` carrega o arquivo a partir da codificação base64 e o anexa ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="a6266-744">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a6266-745">Esse método retorna o identificador de anexo no objeto AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="a6266-745">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a6266-746">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-747">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-747">Parameters:</span></span>
|<span data-ttu-id="a6266-748">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-748">Name</span></span>|<span data-ttu-id="a6266-749">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-749">Type</span></span>|<span data-ttu-id="a6266-750">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-750">Attributes</span></span>|<span data-ttu-id="a6266-751">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-751">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a6266-752">String</span><span class="sxs-lookup"><span data-stu-id="a6266-752">String</span></span>||<span data-ttu-id="a6266-753">O conteúdo codificado em Base 64 de uma imagem ou arquivo a ser adicionado a um email ou evento.</span><span class="sxs-lookup"><span data-stu-id="a6266-753">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a6266-754">String</span><span class="sxs-lookup"><span data-stu-id="a6266-754">String</span></span>||<span data-ttu-id="a6266-p137">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a6266-757">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-757">Object</span></span>|<span data-ttu-id="a6266-758">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-758">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-759">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-759">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-760">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-760">Object</span></span>|<span data-ttu-id="a6266-761">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-761">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-762">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-762">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a6266-763">Booliano</span><span class="sxs-lookup"><span data-stu-id="a6266-763">Boolean</span></span>|<span data-ttu-id="a6266-764">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-764">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-765">Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-765">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a6266-766">function</span><span class="sxs-lookup"><span data-stu-id="a6266-766">function</span></span>|<span data-ttu-id="a6266-767">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-767">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-768">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6266-769">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-769">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a6266-770">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a6266-770">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6266-771">Erros</span><span class="sxs-lookup"><span data-stu-id="a6266-771">Errors</span></span>

|<span data-ttu-id="a6266-772">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a6266-772">Error code</span></span>|<span data-ttu-id="a6266-773">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-773">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a6266-774">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="a6266-774">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a6266-775">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="a6266-775">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a6266-776">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-776">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-777">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-777">Requirements</span></span>

|<span data-ttu-id="a6266-778">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-778">Requirement</span></span>|<span data-ttu-id="a6266-779">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-780">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-781">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-781">Preview</span></span>|
|[<span data-ttu-id="a6266-782">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-782">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-783">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-783">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-784">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-784">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-785">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-785">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6266-786">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a6266-786">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a6266-787">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-787">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a6266-788">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="a6266-788">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a6266-789">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="a6266-789">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-790">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-790">Parameters:</span></span>

| <span data-ttu-id="a6266-791">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-791">Name</span></span> | <span data-ttu-id="a6266-792">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-792">Type</span></span> | <span data-ttu-id="a6266-793">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-793">Attributes</span></span> | <span data-ttu-id="a6266-794">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-794">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a6266-795">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a6266-795">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a6266-796">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a6266-796">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a6266-797">Função</span><span class="sxs-lookup"><span data-stu-id="a6266-797">Function</span></span> || <span data-ttu-id="a6266-p138">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a6266-801">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-801">Object</span></span> | <span data-ttu-id="a6266-802">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-802">&lt;optional&gt;</span></span> | <span data-ttu-id="a6266-803">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-803">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a6266-804">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-804">Object</span></span> | <span data-ttu-id="a6266-805">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-805">&lt;optional&gt;</span></span> | <span data-ttu-id="a6266-806">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-806">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a6266-807">function</span><span class="sxs-lookup"><span data-stu-id="a6266-807">function</span></span>| <span data-ttu-id="a6266-808">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-808">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-809">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-809">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-810">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-810">Requirements</span></span>

|<span data-ttu-id="a6266-811">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-811">Requirement</span></span>| <span data-ttu-id="a6266-812">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-813">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6266-814">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-814">1.7</span></span> |
|[<span data-ttu-id="a6266-815">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-815">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6266-816">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-816">ReadItem</span></span> |
|[<span data-ttu-id="a6266-817">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-817">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6266-818">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-818">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a6266-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a6266-820">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-820">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a6266-p139">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="a6266-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a6266-824">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-824">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a6266-825">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="a6266-825">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-826">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-826">Parameters:</span></span>

|<span data-ttu-id="a6266-827">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-827">Name</span></span>|<span data-ttu-id="a6266-828">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-828">Type</span></span>|<span data-ttu-id="a6266-829">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-829">Attributes</span></span>|<span data-ttu-id="a6266-830">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-830">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a6266-831">String</span><span class="sxs-lookup"><span data-stu-id="a6266-831">String</span></span>||<span data-ttu-id="a6266-p140">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a6266-834">String</span><span class="sxs-lookup"><span data-stu-id="a6266-834">String</span></span>||<span data-ttu-id="a6266-p141">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a6266-837">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-837">Object</span></span>|<span data-ttu-id="a6266-838">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-838">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-839">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-839">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-840">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-840">Object</span></span>|<span data-ttu-id="a6266-841">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-841">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-842">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-842">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-843">function</span><span class="sxs-lookup"><span data-stu-id="a6266-843">function</span></span>|<span data-ttu-id="a6266-844">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-844">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-845">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-845">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6266-846">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-846">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a6266-847">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="a6266-847">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6266-848">Erros</span><span class="sxs-lookup"><span data-stu-id="a6266-848">Errors</span></span>

|<span data-ttu-id="a6266-849">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a6266-849">Error code</span></span>|<span data-ttu-id="a6266-850">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-850">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a6266-851">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-851">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-852">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-852">Requirements</span></span>

|<span data-ttu-id="a6266-853">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-853">Requirement</span></span>|<span data-ttu-id="a6266-854">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-854">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-855">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-855">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-856">1.1</span><span class="sxs-lookup"><span data-stu-id="a6266-856">1.1</span></span>|
|[<span data-ttu-id="a6266-857">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-857">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-858">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-858">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-859">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-859">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-860">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-860">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-861">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-861">Example</span></span>

<span data-ttu-id="a6266-862">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a6266-862">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="a6266-863">close()</span><span class="sxs-lookup"><span data-stu-id="a6266-863">close()</span></span>

<span data-ttu-id="a6266-864">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="a6266-864">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a6266-p142">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="a6266-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-867">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="a6266-867">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a6266-868">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="a6266-868">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-869">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-869">Requirements</span></span>

|<span data-ttu-id="a6266-870">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-870">Requirement</span></span>|<span data-ttu-id="a6266-871">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-872">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-873">1.3</span><span class="sxs-lookup"><span data-stu-id="a6266-873">1.3</span></span>|
|[<span data-ttu-id="a6266-874">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-875">Restrito</span><span class="sxs-lookup"><span data-stu-id="a6266-875">Restricted</span></span>|
|[<span data-ttu-id="a6266-876">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-877">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-877">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a6266-878">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a6266-878">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a6266-879">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a6266-879">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-880">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-880">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-881">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a6266-881">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a6266-882">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a6266-882">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a6266-p143">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a6266-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-886">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-886">Parameters:</span></span>

|<span data-ttu-id="a6266-887">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-887">Name</span></span>|<span data-ttu-id="a6266-888">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-888">Type</span></span>|<span data-ttu-id="a6266-889">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-889">Attributes</span></span>|<span data-ttu-id="a6266-890">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-890">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a6266-891">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a6266-891">String &#124; Object</span></span>||<span data-ttu-id="a6266-p144">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a6266-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a6266-894">**OU**</span><span class="sxs-lookup"><span data-stu-id="a6266-894">**OR**</span></span><br/><span data-ttu-id="a6266-p145">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a6266-897">String</span><span class="sxs-lookup"><span data-stu-id="a6266-897">String</span></span>|<span data-ttu-id="a6266-898">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-898">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a6266-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a6266-901">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-901">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a6266-902">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-902">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-903">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a6266-903">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a6266-904">String</span><span class="sxs-lookup"><span data-stu-id="a6266-904">String</span></span>||<span data-ttu-id="a6266-p147">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a6266-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a6266-907">String</span><span class="sxs-lookup"><span data-stu-id="a6266-907">String</span></span>||<span data-ttu-id="a6266-908">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a6266-908">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a6266-909">String</span><span class="sxs-lookup"><span data-stu-id="a6266-909">String</span></span>||<span data-ttu-id="a6266-p148">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a6266-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a6266-912">Booliano</span><span class="sxs-lookup"><span data-stu-id="a6266-912">Boolean</span></span>||<span data-ttu-id="a6266-p149">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a6266-915">String</span><span class="sxs-lookup"><span data-stu-id="a6266-915">String</span></span>||<span data-ttu-id="a6266-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a6266-919">function</span><span class="sxs-lookup"><span data-stu-id="a6266-919">function</span></span>|<span data-ttu-id="a6266-920">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-920">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-921">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-921">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-922">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-922">Requirements</span></span>

|<span data-ttu-id="a6266-923">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-923">Requirement</span></span>|<span data-ttu-id="a6266-924">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-924">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-925">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-925">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-926">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-926">1.0</span></span>|
|[<span data-ttu-id="a6266-927">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-927">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-928">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-928">ReadItem</span></span>|
|[<span data-ttu-id="a6266-929">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-929">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-930">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-930">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6266-931">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a6266-931">Examples</span></span>

<span data-ttu-id="a6266-932">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a6266-932">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a6266-933">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a6266-933">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a6266-934">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a6266-934">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a6266-935">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a6266-935">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a6266-936">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a6266-936">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a6266-937">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-937">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a6266-938">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a6266-938">displayReplyForm(formData)</span></span>

<span data-ttu-id="a6266-939">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="a6266-939">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-940">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-940">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-941">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="a6266-941">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a6266-942">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a6266-942">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a6266-p151">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="a6266-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-946">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-946">Parameters:</span></span>

|<span data-ttu-id="a6266-947">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-947">Name</span></span>|<span data-ttu-id="a6266-948">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-948">Type</span></span>|<span data-ttu-id="a6266-949">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-949">Attributes</span></span>|<span data-ttu-id="a6266-950">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-950">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a6266-951">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a6266-951">String &#124; Object</span></span>||<span data-ttu-id="a6266-p152">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a6266-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a6266-954">**OU**</span><span class="sxs-lookup"><span data-stu-id="a6266-954">**OR**</span></span><br/><span data-ttu-id="a6266-p153">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a6266-957">String</span><span class="sxs-lookup"><span data-stu-id="a6266-957">String</span></span>|<span data-ttu-id="a6266-958">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-958">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-p154">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a6266-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a6266-961">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-961">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a6266-962">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-962">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-963">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a6266-963">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a6266-964">String</span><span class="sxs-lookup"><span data-stu-id="a6266-964">String</span></span>||<span data-ttu-id="a6266-p155">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a6266-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a6266-967">String</span><span class="sxs-lookup"><span data-stu-id="a6266-967">String</span></span>||<span data-ttu-id="a6266-968">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a6266-968">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a6266-969">String</span><span class="sxs-lookup"><span data-stu-id="a6266-969">String</span></span>||<span data-ttu-id="a6266-p156">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a6266-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a6266-972">Booliano</span><span class="sxs-lookup"><span data-stu-id="a6266-972">Boolean</span></span>||<span data-ttu-id="a6266-p157">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a6266-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a6266-975">String</span><span class="sxs-lookup"><span data-stu-id="a6266-975">String</span></span>||<span data-ttu-id="a6266-p158">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a6266-979">function</span><span class="sxs-lookup"><span data-stu-id="a6266-979">function</span></span>|<span data-ttu-id="a6266-980">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-980">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-981">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-981">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-982">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-982">Requirements</span></span>

|<span data-ttu-id="a6266-983">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-983">Requirement</span></span>|<span data-ttu-id="a6266-984">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-985">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-986">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-986">1.0</span></span>|
|[<span data-ttu-id="a6266-987">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-988">ReadItem</span></span>|
|[<span data-ttu-id="a6266-989">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-990">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-990">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6266-991">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a6266-991">Examples</span></span>

<span data-ttu-id="a6266-992">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a6266-992">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a6266-993">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="a6266-993">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a6266-994">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="a6266-994">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a6266-995">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="a6266-995">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a6266-996">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a6266-996">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a6266-997">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-997">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="a6266-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a6266-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="a6266-999">Obtém o anexo especificado de uma mensagem ou compromisso e o retorna como um objeto `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="a6266-999">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="a6266-1000">O método `getAttachmentContentAsync` remove o obtém anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="a6266-1000">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a6266-1001">Como melhor prática, você deve usar o identificador para recuperar um anexo na mesma sessão da qual attachmentIds foram recuperadas com o chamada `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1001">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="a6266-1002">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-1002">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a6266-1003">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1003">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1004">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1004">Parameters:</span></span>

|<span data-ttu-id="a6266-1005">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1005">Name</span></span>|<span data-ttu-id="a6266-1006">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1006">Type</span></span>|<span data-ttu-id="a6266-1007">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1007">Attributes</span></span>|<span data-ttu-id="a6266-1008">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1008">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a6266-1009">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1009">String</span></span>||<span data-ttu-id="a6266-1010">O identificador do anexo que você quer obter.</span><span class="sxs-lookup"><span data-stu-id="a6266-1010">The identifier of the attachment you want to get.</span></span> <span data-ttu-id="a6266-1011">O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-1011">The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="a6266-1012">Object</span><span class="sxs-lookup"><span data-stu-id="a6266-1012">Object</span></span>|<span data-ttu-id="a6266-1013">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1014">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1015">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1015">Object</span></span>|<span data-ttu-id="a6266-1016">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1017">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1018">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1018">function</span></span>|<span data-ttu-id="a6266-1019">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1020">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1021">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1021">Requirements</span></span>

|<span data-ttu-id="a6266-1022">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1022">Requirement</span></span>|<span data-ttu-id="a6266-1023">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1024">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1025">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-1025">Preview</span></span>|
|[<span data-ttu-id="a6266-1026">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1027">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1028">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1029">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1030">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1030">Returns:</span></span>

<span data-ttu-id="a6266-1031">Tipo: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="a6266-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1032">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1032">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var options = {asyncContext: {type: result.value[i].attachmentType}};
            getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);  
        }
    }
}

function handleAttachmentsCallback(result) {
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a6266-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6266-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a6266-1034">Obtém os anexos do item como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="a6266-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="a6266-1035">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a6266-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1036">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1036">Parameters:</span></span>

|<span data-ttu-id="a6266-1037">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1037">Name</span></span>|<span data-ttu-id="a6266-1038">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1038">Type</span></span>|<span data-ttu-id="a6266-1039">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1039">Attributes</span></span>|<span data-ttu-id="a6266-1040">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a6266-1041">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1041">Object</span></span>|<span data-ttu-id="a6266-1042">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1043">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1044">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1044">Object</span></span>|<span data-ttu-id="a6266-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1046">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1047">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1047">function</span></span>|<span data-ttu-id="a6266-1048">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1049">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1050">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1050">Requirements</span></span>

|<span data-ttu-id="a6266-1051">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1051">Requirement</span></span>|<span data-ttu-id="a6266-1052">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1053">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1054">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-1054">Preview</span></span>|
|[<span data-ttu-id="a6266-1055">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1056">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1057">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1058">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1059">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1059">Returns:</span></span>

<span data-ttu-id="a6266-1060">Tipo: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6266-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1061">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1061">Example</span></span>

<span data-ttu-id="a6266-1062">O exemplo a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a6266-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a6266-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a6266-1064">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1065">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-1066">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1066">Requirements</span></span>

|<span data-ttu-id="a6266-1067">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1067">Requirement</span></span>|<span data-ttu-id="a6266-1068">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1069">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1070">1.0</span></span>|
|[<span data-ttu-id="a6266-1071">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1072">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1073">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1074">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1075">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1075">Returns:</span></span>

<span data-ttu-id="a6266-1076">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a6266-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1077">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1077">Example</span></span>

<span data-ttu-id="a6266-1078">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a6266-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a6266-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a6266-1080">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1081">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1082">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1082">Parameters:</span></span>

|<span data-ttu-id="a6266-1083">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1083">Name</span></span>|<span data-ttu-id="a6266-1084">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1084">Type</span></span>|<span data-ttu-id="a6266-1085">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a6266-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a6266-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a6266-1087">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="a6266-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1088">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1088">Requirements</span></span>

|<span data-ttu-id="a6266-1089">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1089">Requirement</span></span>|<span data-ttu-id="a6266-1090">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1091">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1092">1.0</span></span>|
|[<span data-ttu-id="a6266-1093">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1094">Restrito</span><span class="sxs-lookup"><span data-stu-id="a6266-1094">Restricted</span></span>|
|[<span data-ttu-id="a6266-1095">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1096">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1097">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1097">Returns:</span></span>

<span data-ttu-id="a6266-1098">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="a6266-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a6266-1099">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a6266-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a6266-1100">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a6266-1101">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a6266-1102">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a6266-1102">Value of `entityType`</span></span>|<span data-ttu-id="a6266-1103">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="a6266-1103">Type of objects in returned array</span></span>|<span data-ttu-id="a6266-1104">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="a6266-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a6266-1105">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1105">String</span></span>|<span data-ttu-id="a6266-1106">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a6266-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a6266-1107">Contato</span><span class="sxs-lookup"><span data-stu-id="a6266-1107">Contact</span></span>|<span data-ttu-id="a6266-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6266-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a6266-1109">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1109">String</span></span>|<span data-ttu-id="a6266-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6266-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a6266-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a6266-1111">MeetingSuggestion</span></span>|<span data-ttu-id="a6266-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6266-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a6266-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a6266-1113">PhoneNumber</span></span>|<span data-ttu-id="a6266-1114">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a6266-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a6266-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a6266-1115">TaskSuggestion</span></span>|<span data-ttu-id="a6266-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6266-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a6266-1117">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1117">String</span></span>|<span data-ttu-id="a6266-1118">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="a6266-1118">**Restricted**</span></span>|

<span data-ttu-id="a6266-1119">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a6266-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1120">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1120">Example</span></span>

<span data-ttu-id="a6266-1121">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="a6266-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a6266-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a6266-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a6266-1123">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a6266-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1124">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-1125">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1126">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1126">Parameters:</span></span>

|<span data-ttu-id="a6266-1127">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1127">Name</span></span>|<span data-ttu-id="a6266-1128">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1128">Type</span></span>|<span data-ttu-id="a6266-1129">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a6266-1130">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1130">String</span></span>|<span data-ttu-id="a6266-1131">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a6266-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1132">Requirements</span></span>

|<span data-ttu-id="a6266-1133">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1133">Requirement</span></span>|<span data-ttu-id="a6266-1134">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1136">1.0</span></span>|
|[<span data-ttu-id="a6266-1137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1138">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1140">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1141">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1141">Returns:</span></span>

<span data-ttu-id="a6266-p163">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="a6266-p163">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a6266-1144">Tipo: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a6266-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a6266-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a6266-1146">Obtém dados de inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="a6266-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1147">Esse método só é compatível com o Outlook 2016 ou posterior para Windows (versões Clique para Executar posteriores à 16.0.8413.1000) e o Outlook na Web para o Office 365.</span><span class="sxs-lookup"><span data-stu-id="a6266-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1148">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1148">Parameters:</span></span>
|<span data-ttu-id="a6266-1149">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1149">Name</span></span>|<span data-ttu-id="a6266-1150">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1150">Type</span></span>|<span data-ttu-id="a6266-1151">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1151">Attributes</span></span>|<span data-ttu-id="a6266-1152">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a6266-1153">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1153">Object</span></span>|<span data-ttu-id="a6266-1154">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1155">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1156">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1156">Object</span></span>|<span data-ttu-id="a6266-1157">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1158">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1159">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1159">function</span></span>|<span data-ttu-id="a6266-1160">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1161">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6266-1162">Após o êxito, os dados de inicialização são fornecidos na propriedade `asyncResult.value` como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a6266-1163">Se não houver nenhum contexto de inicialização, o objeto `asyncResult` conterá um objeto `Error` com sua propriedade `code` definida como `9020` e sua propriedade `name` definida como `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1164">Requirements</span></span>

|<span data-ttu-id="a6266-1165">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1165">Requirement</span></span>|<span data-ttu-id="a6266-1166">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1168">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-1168">Preview</span></span>|
|[<span data-ttu-id="a6266-1169">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1170">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1172">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-1173">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1173">Example</span></span>

```javascript
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

#### <a name="getregexmatches--object"></a><span data-ttu-id="a6266-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a6266-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a6266-1175">Retorna valores de cadeia de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a6266-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1176">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-p164">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a6266-p164">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a6266-1180">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a6266-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a6266-1181">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a6266-p165">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="a6266-p165">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-1185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1185">Requirements</span></span>

|<span data-ttu-id="a6266-1186">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1186">Requirement</span></span>|<span data-ttu-id="a6266-1187">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1189">1.0</span></span>|
|[<span data-ttu-id="a6266-1190">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1191">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1192">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1193">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1194">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1194">Returns:</span></span>

<span data-ttu-id="a6266-p166">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a6266-p166">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a6266-1197">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6266-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6266-1198">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6266-1199">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1199">Example</span></span>

<span data-ttu-id="a6266-1200">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="a6266-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a6266-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a6266-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a6266-1202">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a6266-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1203">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-1204">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a6266-p167">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="a6266-p167">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1207">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1207">Parameters:</span></span>

|<span data-ttu-id="a6266-1208">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1208">Name</span></span>|<span data-ttu-id="a6266-1209">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1209">Type</span></span>|<span data-ttu-id="a6266-1210">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a6266-1211">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1211">String</span></span>|<span data-ttu-id="a6266-1212">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="a6266-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1213">Requirements</span></span>

|<span data-ttu-id="a6266-1214">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1214">Requirement</span></span>|<span data-ttu-id="a6266-1215">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1216">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1217">1.0</span></span>|
|[<span data-ttu-id="a6266-1218">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1219">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1221">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1222">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1222">Returns:</span></span>

<span data-ttu-id="a6266-1223">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a6266-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a6266-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6266-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6266-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a6266-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6266-1226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a6266-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a6266-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a6266-1228">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a6266-p168">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p168">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1231">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1231">Parameters:</span></span>

|<span data-ttu-id="a6266-1232">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1232">Name</span></span>|<span data-ttu-id="a6266-1233">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1233">Type</span></span>|<span data-ttu-id="a6266-1234">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1234">Attributes</span></span>|<span data-ttu-id="a6266-1235">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a6266-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a6266-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a6266-p169">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="a6266-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a6266-1240">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1240">Object</span></span>|<span data-ttu-id="a6266-1241">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1242">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1243">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1243">Object</span></span>|<span data-ttu-id="a6266-1244">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1245">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1246">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1246">function</span></span>||<span data-ttu-id="a6266-1247">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6266-1248">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a6266-1249">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1250">Requirements</span></span>

|<span data-ttu-id="a6266-1251">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1251">Requirement</span></span>|<span data-ttu-id="a6266-1252">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="a6266-1254">1.2</span></span>|
|[<span data-ttu-id="a6266-1255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-1257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1258">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1259">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1259">Returns:</span></span>

<span data-ttu-id="a6266-1260">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a6266-1261">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6266-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6266-1262">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6266-1263">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a6266-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a6266-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a6266-p171">Obtém as entidades encontradas em uma correspondência realçada que um usuário selecionou. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a6266-p171">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1267">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-1268">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1268">Requirements</span></span>

|<span data-ttu-id="a6266-1269">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1269">Requirement</span></span>|<span data-ttu-id="a6266-1270">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1271">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="a6266-1272">1.6</span></span>|
|[<span data-ttu-id="a6266-1273">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1274">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1276">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1277">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1277">Returns:</span></span>

<span data-ttu-id="a6266-1278">Tipo: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a6266-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1279">Example</span></span>

<span data-ttu-id="a6266-1280">O exemplo a seguir acessa as entidades de endereços na correspondência realçada, selecionada pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="a6266-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a6266-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a6266-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a6266-p172">Retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML. As correspondências realçadas aplicam-se a [suplementos contextuais](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a6266-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1284">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a6266-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6266-p173">O método `getSelectedRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a6266-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a6266-1288">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="a6266-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a6266-1289">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a6266-p174">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="a6266-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6266-1293">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1293">Requirements</span></span>

|<span data-ttu-id="a6266-1294">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1294">Requirement</span></span>|<span data-ttu-id="a6266-1295">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1296">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="a6266-1297">1.6</span></span>|
|[<span data-ttu-id="a6266-1298">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1299">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1300">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1301">Read</span><span class="sxs-lookup"><span data-stu-id="a6266-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6266-1302">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a6266-1302">Returns:</span></span>

<span data-ttu-id="a6266-p175">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="a6266-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a6266-1305">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1305">Example</span></span>

<span data-ttu-id="a6266-1306">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos de regra de expressão regular `fruits` e `veggies`, que estão especificados no manifesto.</span><span class="sxs-lookup"><span data-stu-id="a6266-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="a6266-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a6266-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="a6266-1308">Obtém as propriedades do compromisso ou mensagem selecionada em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a6266-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1309">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1309">Parameters:</span></span>

|<span data-ttu-id="a6266-1310">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1310">Name</span></span>|<span data-ttu-id="a6266-1311">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1311">Type</span></span>|<span data-ttu-id="a6266-1312">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1312">Attributes</span></span>|<span data-ttu-id="a6266-1313">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a6266-1314">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1314">Object</span></span>|<span data-ttu-id="a6266-1315">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1316">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1317">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1317">Object</span></span>|<span data-ttu-id="a6266-1318">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1319">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1320">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1320">function</span></span>||<span data-ttu-id="a6266-1321">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6266-1322">As propriedades compartilhadas são fornecidas como um objeto [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a6266-1323">Esse objeto pode ser usado para obter as propriedades compartilhadas do item.</span><span class="sxs-lookup"><span data-stu-id="a6266-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1324">Requirements</span></span>

|<span data-ttu-id="a6266-1325">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1325">Requirement</span></span>|<span data-ttu-id="a6266-1326">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1328">Visualização</span><span class="sxs-lookup"><span data-stu-id="a6266-1328">Preview</span></span>|
|[<span data-ttu-id="a6266-1329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1330">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1332">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-1333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a6266-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a6266-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a6266-1335">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a6266-p177">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="a6266-p177">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1339">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1339">Parameters:</span></span>

|<span data-ttu-id="a6266-1340">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1340">Name</span></span>|<span data-ttu-id="a6266-1341">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1341">Type</span></span>|<span data-ttu-id="a6266-1342">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1342">Attributes</span></span>|<span data-ttu-id="a6266-1343">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a6266-1344">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1344">function</span></span>||<span data-ttu-id="a6266-1345">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6266-1346">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a6266-1347">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="a6266-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a6266-1348">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1348">Object</span></span>|<span data-ttu-id="a6266-1349">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1350">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a6266-1351">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1352">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1352">Requirements</span></span>

|<span data-ttu-id="a6266-1353">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1353">Requirement</span></span>|<span data-ttu-id="a6266-1354">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1355">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="a6266-1356">1.0</span></span>|
|[<span data-ttu-id="a6266-1357">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1358">ReadItem</span></span>|
|[<span data-ttu-id="a6266-1359">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1360">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-1361">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1361">Example</span></span>

<span data-ttu-id="a6266-p180">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a6266-p180">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a6266-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6266-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a6266-1366">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="a6266-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a6266-1367">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item.</span><span class="sxs-lookup"><span data-stu-id="a6266-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a6266-1368">Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a6266-1369">No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="a6266-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a6266-1370">Uma sessão é finalizada quando o usuário fecha o aplicativo, ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1371">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1371">Parameters:</span></span>

|<span data-ttu-id="a6266-1372">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1372">Name</span></span>|<span data-ttu-id="a6266-1373">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1373">Type</span></span>|<span data-ttu-id="a6266-1374">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1374">Attributes</span></span>|<span data-ttu-id="a6266-1375">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a6266-1376">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1376">String</span></span>||<span data-ttu-id="a6266-p182">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a6266-p182">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="a6266-1379">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1379">Object</span></span>|<span data-ttu-id="a6266-1380">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1381">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1382">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1382">Object</span></span>|<span data-ttu-id="a6266-1383">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1384">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1385">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1385">function</span></span>|<span data-ttu-id="a6266-1386">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1387">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1387">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6266-1388">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="a6266-1388">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6266-1389">Erros</span><span class="sxs-lookup"><span data-stu-id="a6266-1389">Errors</span></span>

|<span data-ttu-id="a6266-1390">Código de erro</span><span class="sxs-lookup"><span data-stu-id="a6266-1390">Error code</span></span>|<span data-ttu-id="a6266-1391">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1391">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a6266-1392">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="a6266-1392">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1393">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1393">Requirements</span></span>

|<span data-ttu-id="a6266-1394">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1394">Requirement</span></span>|<span data-ttu-id="a6266-1395">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1395">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1396">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1397">1.1</span><span class="sxs-lookup"><span data-stu-id="a6266-1397">1.1</span></span>|
|[<span data-ttu-id="a6266-1398">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1398">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1399">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1399">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-1400">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1400">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1401">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-1401">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-1402">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1402">Example</span></span>

<span data-ttu-id="a6266-1403">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="a6266-1403">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a6266-1404">removeHandlerAsync(eventType, handler, [opções], [retorno de chamada])</span><span class="sxs-lookup"><span data-stu-id="a6266-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a6266-1405">Remove um manipulador de eventos de um evento compatível.</span><span class="sxs-lookup"><span data-stu-id="a6266-1405">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="a6266-1406">Atualmente, os tipos de evento compatíveis são `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` e `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1406">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1407">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1407">Parameters:</span></span>

| <span data-ttu-id="a6266-1408">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1408">Name</span></span> | <span data-ttu-id="a6266-1409">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1409">Type</span></span> | <span data-ttu-id="a6266-1410">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1410">Attributes</span></span> | <span data-ttu-id="a6266-1411">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1411">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a6266-1412">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a6266-1412">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a6266-1413">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a6266-1413">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a6266-1414">Função</span><span class="sxs-lookup"><span data-stu-id="a6266-1414">Function</span></span> || <span data-ttu-id="a6266-p183">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p183">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a6266-1418">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1418">Object</span></span> | <span data-ttu-id="a6266-1419">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1419">&lt;optional&gt;</span></span> | <span data-ttu-id="a6266-1420">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1420">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a6266-1421">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1421">Object</span></span> | <span data-ttu-id="a6266-1422">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1422">&lt;optional&gt;</span></span> | <span data-ttu-id="a6266-1423">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1423">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a6266-1424">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1424">function</span></span>| <span data-ttu-id="a6266-1425">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1425">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1426">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1427">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1427">Requirements</span></span>

|<span data-ttu-id="a6266-1428">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1428">Requirement</span></span>| <span data-ttu-id="a6266-1429">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1429">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1430">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6266-1431">1.7</span><span class="sxs-lookup"><span data-stu-id="a6266-1431">1.7</span></span> |
|[<span data-ttu-id="a6266-1432">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6266-1433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1433">ReadItem</span></span> |
|[<span data-ttu-id="a6266-1434">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6266-1435">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="a6266-1435">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a6266-1436">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a6266-1436">saveAsync([options], callback)</span></span>

<span data-ttu-id="a6266-1437">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a6266-1437">Asynchronously saves an item.</span></span>

<span data-ttu-id="a6266-p184">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="a6266-p184">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1441">Se seu suplemento chamar o `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="a6266-1441">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a6266-1442">Até que o item esteja sincronizado, usar o `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="a6266-1442">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a6266-p186">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="a6266-p186">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a6266-1446">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="a6266-1446">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a6266-1447">O Outlook para Mac não dá suporte ao `saveAsync` em uma reunião no modo composto.</span><span class="sxs-lookup"><span data-stu-id="a6266-1447">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="a6266-1448">Chamar `saveAsync` em uma reunião no Outlook para Mac fará com que um erro seja retornado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1448">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a6266-1449">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="a6266-1449">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1450">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1450">Parameters:</span></span>

|<span data-ttu-id="a6266-1451">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1451">Name</span></span>|<span data-ttu-id="a6266-1452">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1452">Type</span></span>|<span data-ttu-id="a6266-1453">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1453">Attributes</span></span>|<span data-ttu-id="a6266-1454">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1454">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a6266-1455">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1455">Object</span></span>|<span data-ttu-id="a6266-1456">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1456">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1457">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1457">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1458">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1458">Object</span></span>|<span data-ttu-id="a6266-1459">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1459">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1460">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1460">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a6266-1461">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1461">function</span></span>||<span data-ttu-id="a6266-1462">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1462">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6266-1463">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6266-1463">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1464">Requirements</span></span>

|<span data-ttu-id="a6266-1465">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1465">Requirement</span></span>|<span data-ttu-id="a6266-1466">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1468">1.3</span><span class="sxs-lookup"><span data-stu-id="a6266-1468">1.3</span></span>|
|[<span data-ttu-id="a6266-1469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1470">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1470">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-1471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1472">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-1472">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6266-1473">Exemplos</span><span class="sxs-lookup"><span data-stu-id="a6266-1473">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a6266-p188">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="a6266-p188">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a6266-1476">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a6266-1476">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a6266-1477">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6266-1477">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a6266-p189">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="a6266-p189">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6266-1481">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a6266-1481">Parameters:</span></span>

|<span data-ttu-id="a6266-1482">Nome</span><span class="sxs-lookup"><span data-stu-id="a6266-1482">Name</span></span>|<span data-ttu-id="a6266-1483">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6266-1483">Type</span></span>|<span data-ttu-id="a6266-1484">Atributos</span><span class="sxs-lookup"><span data-stu-id="a6266-1484">Attributes</span></span>|<span data-ttu-id="a6266-1485">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6266-1485">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a6266-1486">String</span><span class="sxs-lookup"><span data-stu-id="a6266-1486">String</span></span>||<span data-ttu-id="a6266-p190">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p190">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a6266-1490">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1490">Object</span></span>|<span data-ttu-id="a6266-1491">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1492">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a6266-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a6266-1493">Objeto</span><span class="sxs-lookup"><span data-stu-id="a6266-1493">Object</span></span>|<span data-ttu-id="a6266-1494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-1495">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a6266-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a6266-1496">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a6266-1496">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a6266-1497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6266-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="a6266-p191">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="a6266-p191">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a6266-p192">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a6266-p192">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a6266-1502">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="a6266-1502">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a6266-1503">function</span><span class="sxs-lookup"><span data-stu-id="a6266-1503">function</span></span>||<span data-ttu-id="a6266-1504">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6266-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6266-1505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6266-1505">Requirements</span></span>

|<span data-ttu-id="a6266-1506">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6266-1506">Requirement</span></span>|<span data-ttu-id="a6266-1507">Valor</span><span class="sxs-lookup"><span data-stu-id="a6266-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6266-1508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6266-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a6266-1509">1.2</span><span class="sxs-lookup"><span data-stu-id="a6266-1509">1.2</span></span>|
|[<span data-ttu-id="a6266-1510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a6266-1510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a6266-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6266-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6266-1512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6266-1512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a6266-1513">Escrever</span><span class="sxs-lookup"><span data-stu-id="a6266-1513">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6266-1514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a6266-1514">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```