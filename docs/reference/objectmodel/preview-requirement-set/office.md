 

# <a name="office"></a><span data-ttu-id="ec7a8-101">Office</span><span class="sxs-lookup"><span data-stu-id="ec7a8-101">Office</span></span>

<span data-ttu-id="ec7a8-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ec7a8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec7a8-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-104">Requirements</span></span>

|<span data-ttu-id="ec7a8-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec7a8-105">Requirement</span></span>| <span data-ttu-id="ec7a8-106">Valor</span><span class="sxs-lookup"><span data-stu-id="ec7a8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec7a8-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec7a8-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec7a8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ec7a8-108">1.0</span></span>|
|[<span data-ttu-id="ec7a8-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec7a8-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec7a8-110">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ec7a8-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ec7a8-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-111">Members and methods</span></span>

| <span data-ttu-id="ec7a8-112">Membro</span><span class="sxs-lookup"><span data-stu-id="ec7a8-112">Member</span></span> | <span data-ttu-id="ec7a8-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec7a8-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ec7a8-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ec7a8-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ec7a8-115">Membro</span><span class="sxs-lookup"><span data-stu-id="ec7a8-115">Member</span></span> |
| [<span data-ttu-id="ec7a8-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ec7a8-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ec7a8-117">Membro</span><span class="sxs-lookup"><span data-stu-id="ec7a8-117">Member</span></span> |
| [<span data-ttu-id="ec7a8-118">EventType</span><span class="sxs-lookup"><span data-stu-id="ec7a8-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ec7a8-119">Membro</span><span class="sxs-lookup"><span data-stu-id="ec7a8-119">Member</span></span> |
| [<span data-ttu-id="ec7a8-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ec7a8-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ec7a8-121">Membro</span><span class="sxs-lookup"><span data-stu-id="ec7a8-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ec7a8-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ec7a8-122">Namespaces</span></span>

<span data-ttu-id="ec7a8-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ec7a8-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ec7a8-125">Membros</span><span class="sxs-lookup"><span data-stu-id="ec7a8-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ec7a8-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="ec7a8-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ec7a8-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-128">Type:</span></span>

*   <span data-ttu-id="ec7a8-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec7a8-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-130">Properties:</span></span>

|<span data-ttu-id="ec7a8-131">Nome</span><span class="sxs-lookup"><span data-stu-id="ec7a8-131">Name</span></span>| <span data-ttu-id="ec7a8-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec7a8-132">Type</span></span>| <span data-ttu-id="ec7a8-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec7a8-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ec7a8-134">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-134">String</span></span>|<span data-ttu-id="ec7a8-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ec7a8-136">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-136">String</span></span>|<span data-ttu-id="ec7a8-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec7a8-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-138">Requirements</span></span>

|<span data-ttu-id="ec7a8-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec7a8-139">Requirement</span></span>| <span data-ttu-id="ec7a8-140">Valor</span><span class="sxs-lookup"><span data-stu-id="ec7a8-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec7a8-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec7a8-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec7a8-142">1.0</span><span class="sxs-lookup"><span data-stu-id="ec7a8-142">1.0</span></span>|
|[<span data-ttu-id="ec7a8-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec7a8-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec7a8-144">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ec7a8-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ec7a8-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-145">CoercionType :String</span></span>

<span data-ttu-id="ec7a8-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ec7a8-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-147">Type:</span></span>

*   <span data-ttu-id="ec7a8-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec7a8-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-149">Properties:</span></span>

|<span data-ttu-id="ec7a8-150">Nome</span><span class="sxs-lookup"><span data-stu-id="ec7a8-150">Name</span></span>| <span data-ttu-id="ec7a8-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec7a8-151">Type</span></span>| <span data-ttu-id="ec7a8-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec7a8-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ec7a8-153">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-153">String</span></span>|<span data-ttu-id="ec7a8-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ec7a8-155">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-155">String</span></span>|<span data-ttu-id="ec7a8-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec7a8-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-157">Requirements</span></span>

|<span data-ttu-id="ec7a8-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec7a8-158">Requirement</span></span>| <span data-ttu-id="ec7a8-159">Valor</span><span class="sxs-lookup"><span data-stu-id="ec7a8-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec7a8-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec7a8-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec7a8-161">1.0</span><span class="sxs-lookup"><span data-stu-id="ec7a8-161">1.0</span></span>|
|[<span data-ttu-id="ec7a8-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec7a8-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec7a8-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ec7a8-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ec7a8-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-164">EventType :String</span></span>

<span data-ttu-id="ec7a8-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ec7a8-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-166">Type:</span></span>

*   <span data-ttu-id="ec7a8-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec7a8-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-168">Properties:</span></span>

| <span data-ttu-id="ec7a8-169">Nome</span><span class="sxs-lookup"><span data-stu-id="ec7a8-169">Name</span></span> | <span data-ttu-id="ec7a8-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec7a8-170">Type</span></span> | <span data-ttu-id="ec7a8-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec7a8-171">Description</span></span> | <span data-ttu-id="ec7a8-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="ec7a8-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-173">String</span></span> | <span data-ttu-id="ec7a8-174">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ec7a8-175">1.7</span><span class="sxs-lookup"><span data-stu-id="ec7a8-175">-17</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="ec7a8-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-176">String</span></span> | <span data-ttu-id="ec7a8-177">Um anexo foi adicionado a ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-177">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="ec7a8-178">Visualização</span><span class="sxs-lookup"><span data-stu-id="ec7a8-178">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="ec7a8-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-179">String</span></span> | <span data-ttu-id="ec7a8-180">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-180">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ec7a8-181">1.5</span><span class="sxs-lookup"><span data-stu-id="ec7a8-181">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="ec7a8-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-182">String</span></span> | <span data-ttu-id="ec7a8-183">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-183">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="ec7a8-184">Visualização</span><span class="sxs-lookup"><span data-stu-id="ec7a8-184">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ec7a8-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-185">String</span></span> | <span data-ttu-id="ec7a8-186">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-186">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ec7a8-187">1.7</span><span class="sxs-lookup"><span data-stu-id="ec7a8-187">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ec7a8-188">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-188">String</span></span> | <span data-ttu-id="ec7a8-189">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-189">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ec7a8-190">1.7</span><span class="sxs-lookup"><span data-stu-id="ec7a8-190">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec7a8-191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-191">Requirements</span></span>

|<span data-ttu-id="ec7a8-192">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec7a8-192">Requirement</span></span>| <span data-ttu-id="ec7a8-193">Valor</span><span class="sxs-lookup"><span data-stu-id="ec7a8-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec7a8-194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec7a8-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec7a8-195">1.5</span><span class="sxs-lookup"><span data-stu-id="ec7a8-195">1.5</span></span> |
|[<span data-ttu-id="ec7a8-196">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec7a8-196">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec7a8-197">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ec7a8-197">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ec7a8-198">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-198">SourceProperty :String</span></span>

<span data-ttu-id="ec7a8-199">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-199">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ec7a8-200">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-200">Type:</span></span>

*   <span data-ttu-id="ec7a8-201">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec7a8-201">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec7a8-202">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec7a8-202">Properties:</span></span>

|<span data-ttu-id="ec7a8-203">Nome</span><span class="sxs-lookup"><span data-stu-id="ec7a8-203">Name</span></span>| <span data-ttu-id="ec7a8-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec7a8-204">Type</span></span>| <span data-ttu-id="ec7a8-205">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec7a8-205">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ec7a8-206">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-206">String</span></span>|<span data-ttu-id="ec7a8-207">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-207">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ec7a8-208">String</span><span class="sxs-lookup"><span data-stu-id="ec7a8-208">String</span></span>|<span data-ttu-id="ec7a8-209">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ec7a8-209">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec7a8-210">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec7a8-210">Requirements</span></span>

|<span data-ttu-id="ec7a8-211">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec7a8-211">Requirement</span></span>| <span data-ttu-id="ec7a8-212">Valor</span><span class="sxs-lookup"><span data-stu-id="ec7a8-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec7a8-213">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec7a8-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec7a8-214">1.0</span><span class="sxs-lookup"><span data-stu-id="ec7a8-214">1.0</span></span>|
|[<span data-ttu-id="ec7a8-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec7a8-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec7a8-216">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ec7a8-216">Compose or read</span></span>|