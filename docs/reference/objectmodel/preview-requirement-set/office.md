 

# <a name="office"></a><span data-ttu-id="243e2-101">Office</span><span class="sxs-lookup"><span data-stu-id="243e2-101">Office</span></span>

<span data-ttu-id="243e2-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="243e2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="243e2-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="243e2-104">Requirements</span></span>

|<span data-ttu-id="243e2-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="243e2-105">Requirement</span></span>| <span data-ttu-id="243e2-106">Valor</span><span class="sxs-lookup"><span data-stu-id="243e2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="243e2-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="243e2-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243e2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="243e2-108">1.0</span></span>|
|[<span data-ttu-id="243e2-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="243e2-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="243e2-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="243e2-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="243e2-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="243e2-111">Members and methods</span></span>

| <span data-ttu-id="243e2-112">Membro</span><span class="sxs-lookup"><span data-stu-id="243e2-112">Member</span></span> | <span data-ttu-id="243e2-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="243e2-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="243e2-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="243e2-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="243e2-115">Membro</span><span class="sxs-lookup"><span data-stu-id="243e2-115">Member</span></span> |
| [<span data-ttu-id="243e2-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="243e2-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="243e2-117">Membro</span><span class="sxs-lookup"><span data-stu-id="243e2-117">Member</span></span> |
| [<span data-ttu-id="243e2-118">EventType</span><span class="sxs-lookup"><span data-stu-id="243e2-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="243e2-119">Membro</span><span class="sxs-lookup"><span data-stu-id="243e2-119">Member</span></span> |
| [<span data-ttu-id="243e2-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="243e2-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="243e2-121">Membro</span><span class="sxs-lookup"><span data-stu-id="243e2-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="243e2-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="243e2-122">Namespaces</span></span>

<span data-ttu-id="243e2-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API dos suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="243e2-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="243e2-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="243e2-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="243e2-125">Membros</span><span class="sxs-lookup"><span data-stu-id="243e2-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="243e2-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="243e2-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="243e2-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="243e2-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="243e2-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="243e2-128">Type:</span></span>

*   <span data-ttu-id="243e2-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243e2-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="243e2-130">Properties:</span></span>

|<span data-ttu-id="243e2-131">Nome</span><span class="sxs-lookup"><span data-stu-id="243e2-131">Name</span></span>| <span data-ttu-id="243e2-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="243e2-132">Type</span></span>| <span data-ttu-id="243e2-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="243e2-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="243e2-134">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-134">String</span></span>|<span data-ttu-id="243e2-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="243e2-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="243e2-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-136">String</span></span>|<span data-ttu-id="243e2-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="243e2-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243e2-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="243e2-138">Requirements</span></span>

|<span data-ttu-id="243e2-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="243e2-139">Requirement</span></span>| <span data-ttu-id="243e2-140">Valor</span><span class="sxs-lookup"><span data-stu-id="243e2-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="243e2-141">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="243e2-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243e2-142">1.0</span><span class="sxs-lookup"><span data-stu-id="243e2-142">1.0</span></span>|
|[<span data-ttu-id="243e2-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="243e2-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="243e2-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="243e2-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="243e2-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="243e2-145">CoercionType :String</span></span>

<span data-ttu-id="243e2-146">Especifica como forçar os dados retornados ou definir de acordo com o método invocado.</span><span class="sxs-lookup"><span data-stu-id="243e2-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243e2-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="243e2-147">Type:</span></span>

*   <span data-ttu-id="243e2-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243e2-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="243e2-149">Properties:</span></span>

|<span data-ttu-id="243e2-150">Nome</span><span class="sxs-lookup"><span data-stu-id="243e2-150">Name</span></span>| <span data-ttu-id="243e2-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="243e2-151">Type</span></span>| <span data-ttu-id="243e2-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="243e2-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="243e2-153">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-153">String</span></span>|<span data-ttu-id="243e2-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="243e2-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="243e2-155">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-155">String</span></span>|<span data-ttu-id="243e2-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="243e2-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243e2-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="243e2-157">Requirements</span></span>

|<span data-ttu-id="243e2-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="243e2-158">Requirement</span></span>| <span data-ttu-id="243e2-159">Valor</span><span class="sxs-lookup"><span data-stu-id="243e2-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="243e2-160">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="243e2-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243e2-161">1.0</span><span class="sxs-lookup"><span data-stu-id="243e2-161">1.0</span></span>|
|[<span data-ttu-id="243e2-162">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="243e2-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="243e2-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="243e2-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="243e2-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="243e2-164">EventType :String</span></span>

<span data-ttu-id="243e2-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="243e2-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="243e2-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="243e2-166">Type:</span></span>

*   <span data-ttu-id="243e2-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243e2-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="243e2-168">Properties:</span></span>

| <span data-ttu-id="243e2-169">Nome</span><span class="sxs-lookup"><span data-stu-id="243e2-169">Name</span></span> | <span data-ttu-id="243e2-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="243e2-170">Type</span></span> | <span data-ttu-id="243e2-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="243e2-171">Description</span></span> | <span data-ttu-id="243e2-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="243e2-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="243e2-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-173">String</span></span> | <span data-ttu-id="243e2-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="243e2-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="243e2-175">1.7</span><span class="sxs-lookup"><span data-stu-id="243e2-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="243e2-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-176">String</span></span> | <span data-ttu-id="243e2-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="243e2-177">The selected item has changed.</span></span> | <span data-ttu-id="243e2-178">1.5</span><span class="sxs-lookup"><span data-stu-id="243e2-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="243e2-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-179">String</span></span> | <span data-ttu-id="243e2-180">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="243e2-180">The selected item has changed.</span></span> | <span data-ttu-id="243e2-181">Visualizar</span><span class="sxs-lookup"><span data-stu-id="243e2-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="243e2-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-182">String</span></span> | <span data-ttu-id="243e2-183">A lista de destinatários do item selecionado ou o local do compromisso foram alterados.</span><span class="sxs-lookup"><span data-stu-id="243e2-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="243e2-184">1.7</span><span class="sxs-lookup"><span data-stu-id="243e2-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="243e2-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-185">String</span></span> | <span data-ttu-id="243e2-186">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="243e2-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="243e2-187">1.7</span><span class="sxs-lookup"><span data-stu-id="243e2-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="243e2-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="243e2-188">Requirements</span></span>

|<span data-ttu-id="243e2-189">Requisito</span><span class="sxs-lookup"><span data-stu-id="243e2-189">Requirement</span></span>| <span data-ttu-id="243e2-190">Valor</span><span class="sxs-lookup"><span data-stu-id="243e2-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="243e2-191">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="243e2-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243e2-192">1.5</span><span class="sxs-lookup"><span data-stu-id="243e2-192">1.5</span></span> |
|[<span data-ttu-id="243e2-193">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="243e2-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="243e2-194">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="243e2-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="243e2-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="243e2-195">SourceProperty :String</span></span>

<span data-ttu-id="243e2-196">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="243e2-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243e2-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="243e2-197">Type:</span></span>

*   <span data-ttu-id="243e2-198">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243e2-199">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="243e2-199">Properties:</span></span>

|<span data-ttu-id="243e2-200">Nome</span><span class="sxs-lookup"><span data-stu-id="243e2-200">Name</span></span>| <span data-ttu-id="243e2-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="243e2-201">Type</span></span>| <span data-ttu-id="243e2-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="243e2-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="243e2-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-203">String</span></span>|<span data-ttu-id="243e2-204">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="243e2-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="243e2-205">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="243e2-205">String</span></span>|<span data-ttu-id="243e2-206">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="243e2-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243e2-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="243e2-207">Requirements</span></span>

|<span data-ttu-id="243e2-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="243e2-208">Requirement</span></span>| <span data-ttu-id="243e2-209">Valor</span><span class="sxs-lookup"><span data-stu-id="243e2-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="243e2-210">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="243e2-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="243e2-211">1.0</span><span class="sxs-lookup"><span data-stu-id="243e2-211">1.0</span></span>|
|[<span data-ttu-id="243e2-212">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="243e2-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="243e2-213">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="243e2-213">Compose or read</span></span>|