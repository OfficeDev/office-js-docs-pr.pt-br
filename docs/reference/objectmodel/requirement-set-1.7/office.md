 

# <a name="office"></a><span data-ttu-id="bbcd0-101">Office</span><span class="sxs-lookup"><span data-stu-id="bbcd0-101">Office</span></span>

<span data-ttu-id="bbcd0-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="bbcd0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bbcd0-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-104">Requirements</span></span>

|<span data-ttu-id="bbcd0-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="bbcd0-105">Requirement</span></span>| <span data-ttu-id="bbcd0-106">Valor</span><span class="sxs-lookup"><span data-stu-id="bbcd0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbcd0-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bbcd0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbcd0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="bbcd0-108">1.0</span></span>|
|[<span data-ttu-id="bbcd0-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bbcd0-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbcd0-110">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="bbcd0-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bbcd0-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-111">Members and methods</span></span>

| <span data-ttu-id="bbcd0-112">Membro</span><span class="sxs-lookup"><span data-stu-id="bbcd0-112">Member</span></span> | <span data-ttu-id="bbcd0-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="bbcd0-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bbcd0-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bbcd0-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bbcd0-115">Membro</span><span class="sxs-lookup"><span data-stu-id="bbcd0-115">Member</span></span> |
| [<span data-ttu-id="bbcd0-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bbcd0-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bbcd0-117">Membro</span><span class="sxs-lookup"><span data-stu-id="bbcd0-117">Member</span></span> |
| [<span data-ttu-id="bbcd0-118">EventType</span><span class="sxs-lookup"><span data-stu-id="bbcd0-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="bbcd0-119">Membro</span><span class="sxs-lookup"><span data-stu-id="bbcd0-119">Member</span></span> |
| [<span data-ttu-id="bbcd0-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bbcd0-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bbcd0-121">Membro</span><span class="sxs-lookup"><span data-stu-id="bbcd0-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="bbcd0-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="bbcd0-122">Namespaces</span></span>

<span data-ttu-id="bbcd0-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="bbcd0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="bbcd0-125">Membros</span><span class="sxs-lookup"><span data-stu-id="bbcd0-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="bbcd0-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="bbcd0-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bbcd0-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-128">Type:</span></span>

*   <span data-ttu-id="bbcd0-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bbcd0-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-130">Properties:</span></span>

|<span data-ttu-id="bbcd0-131">Nome</span><span class="sxs-lookup"><span data-stu-id="bbcd0-131">Name</span></span>| <span data-ttu-id="bbcd0-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="bbcd0-132">Type</span></span>| <span data-ttu-id="bbcd0-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="bbcd0-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bbcd0-134">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-134">String</span></span>|<span data-ttu-id="bbcd0-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bbcd0-136">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-136">String</span></span>|<span data-ttu-id="bbcd0-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bbcd0-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-138">Requirements</span></span>

|<span data-ttu-id="bbcd0-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="bbcd0-139">Requirement</span></span>| <span data-ttu-id="bbcd0-140">Valor</span><span class="sxs-lookup"><span data-stu-id="bbcd0-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbcd0-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bbcd0-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbcd0-142">1.0</span><span class="sxs-lookup"><span data-stu-id="bbcd0-142">1.0</span></span>|
|[<span data-ttu-id="bbcd0-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bbcd0-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbcd0-144">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="bbcd0-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="bbcd0-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-145">CoercionType :String</span></span>

<span data-ttu-id="bbcd0-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bbcd0-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-147">Type:</span></span>

*   <span data-ttu-id="bbcd0-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bbcd0-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-149">Properties:</span></span>

|<span data-ttu-id="bbcd0-150">Nome</span><span class="sxs-lookup"><span data-stu-id="bbcd0-150">Name</span></span>| <span data-ttu-id="bbcd0-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="bbcd0-151">Type</span></span>| <span data-ttu-id="bbcd0-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="bbcd0-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bbcd0-153">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-153">String</span></span>|<span data-ttu-id="bbcd0-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bbcd0-155">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-155">String</span></span>|<span data-ttu-id="bbcd0-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bbcd0-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-157">Requirements</span></span>

|<span data-ttu-id="bbcd0-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="bbcd0-158">Requirement</span></span>| <span data-ttu-id="bbcd0-159">Valor</span><span class="sxs-lookup"><span data-stu-id="bbcd0-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbcd0-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bbcd0-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbcd0-161">1.0</span><span class="sxs-lookup"><span data-stu-id="bbcd0-161">1.0</span></span>|
|[<span data-ttu-id="bbcd0-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bbcd0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbcd0-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="bbcd0-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="bbcd0-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-164">EventType :String</span></span>

<span data-ttu-id="bbcd0-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="bbcd0-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-166">Type:</span></span>

*   <span data-ttu-id="bbcd0-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bbcd0-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-168">Properties:</span></span>

| <span data-ttu-id="bbcd0-169">Nome</span><span class="sxs-lookup"><span data-stu-id="bbcd0-169">Name</span></span> | <span data-ttu-id="bbcd0-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="bbcd0-170">Type</span></span> | <span data-ttu-id="bbcd0-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="bbcd0-171">Description</span></span> | <span data-ttu-id="bbcd0-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="bbcd0-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-173">String</span></span> | <span data-ttu-id="bbcd0-174">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="bbcd0-175">1.7</span><span class="sxs-lookup"><span data-stu-id="bbcd0-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="bbcd0-176">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-176">String</span></span> | <span data-ttu-id="bbcd0-177">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-177">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="bbcd0-178">1.5</span><span class="sxs-lookup"><span data-stu-id="bbcd0-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="bbcd0-179">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-179">String</span></span> | <span data-ttu-id="bbcd0-180">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="bbcd0-181">1.7</span><span class="sxs-lookup"><span data-stu-id="bbcd0-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="bbcd0-182">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-182">String</span></span> | <span data-ttu-id="bbcd0-183">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="bbcd0-184">1.7</span><span class="sxs-lookup"><span data-stu-id="bbcd0-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bbcd0-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-185">Requirements</span></span>

|<span data-ttu-id="bbcd0-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="bbcd0-186">Requirement</span></span>| <span data-ttu-id="bbcd0-187">Valor</span><span class="sxs-lookup"><span data-stu-id="bbcd0-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbcd0-188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bbcd0-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbcd0-189">1.5</span><span class="sxs-lookup"><span data-stu-id="bbcd0-189">1.5</span></span> |
|[<span data-ttu-id="bbcd0-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bbcd0-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbcd0-191">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="bbcd0-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="bbcd0-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-192">SourceProperty :String</span></span>

<span data-ttu-id="bbcd0-193">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bbcd0-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-194">Type:</span></span>

*   <span data-ttu-id="bbcd0-195">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bbcd0-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bbcd0-196">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="bbcd0-196">Properties:</span></span>

|<span data-ttu-id="bbcd0-197">Nome</span><span class="sxs-lookup"><span data-stu-id="bbcd0-197">Name</span></span>| <span data-ttu-id="bbcd0-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="bbcd0-198">Type</span></span>| <span data-ttu-id="bbcd0-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="bbcd0-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bbcd0-200">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-200">String</span></span>|<span data-ttu-id="bbcd0-201">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bbcd0-202">String</span><span class="sxs-lookup"><span data-stu-id="bbcd0-202">String</span></span>|<span data-ttu-id="bbcd0-203">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="bbcd0-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bbcd0-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bbcd0-204">Requirements</span></span>

|<span data-ttu-id="bbcd0-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="bbcd0-205">Requirement</span></span>| <span data-ttu-id="bbcd0-206">Valor</span><span class="sxs-lookup"><span data-stu-id="bbcd0-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="bbcd0-207">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bbcd0-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bbcd0-208">1.0</span><span class="sxs-lookup"><span data-stu-id="bbcd0-208">1.0</span></span>|
|[<span data-ttu-id="bbcd0-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bbcd0-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bbcd0-210">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="bbcd0-210">Compose or read</span></span>|