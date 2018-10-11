 

# <a name="office"></a><span data-ttu-id="55cc6-101">Office</span><span class="sxs-lookup"><span data-stu-id="55cc6-101">Office</span></span>

<span data-ttu-id="55cc6-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="55cc6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="55cc6-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55cc6-104">Requirements</span></span>

|<span data-ttu-id="55cc6-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="55cc6-105">Requirement</span></span>| <span data-ttu-id="55cc6-106">Valor</span><span class="sxs-lookup"><span data-stu-id="55cc6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="55cc6-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55cc6-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55cc6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="55cc6-108">1.0</span></span>|
|[<span data-ttu-id="55cc6-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55cc6-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55cc6-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="55cc6-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="55cc6-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="55cc6-111">Members and methods</span></span>

| <span data-ttu-id="55cc6-112">Membro</span><span class="sxs-lookup"><span data-stu-id="55cc6-112">Member</span></span> | <span data-ttu-id="55cc6-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="55cc6-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="55cc6-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="55cc6-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="55cc6-115">Membro</span><span class="sxs-lookup"><span data-stu-id="55cc6-115">Member</span></span> |
| [<span data-ttu-id="55cc6-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="55cc6-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="55cc6-117">Membro</span><span class="sxs-lookup"><span data-stu-id="55cc6-117">Member</span></span> |
| [<span data-ttu-id="55cc6-118">EventType</span><span class="sxs-lookup"><span data-stu-id="55cc6-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="55cc6-119">Membro</span><span class="sxs-lookup"><span data-stu-id="55cc6-119">Member</span></span> |
| [<span data-ttu-id="55cc6-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="55cc6-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="55cc6-121">Membro</span><span class="sxs-lookup"><span data-stu-id="55cc6-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="55cc6-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="55cc6-122">Namespaces</span></span>

<span data-ttu-id="55cc6-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="55cc6-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="55cc6-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="55cc6-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="55cc6-125">Membros</span><span class="sxs-lookup"><span data-stu-id="55cc6-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="55cc6-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="55cc6-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="55cc6-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="55cc6-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="55cc6-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="55cc6-128">Type:</span></span>

*   <span data-ttu-id="55cc6-129">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55cc6-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55cc6-130">Properties:</span></span>

|<span data-ttu-id="55cc6-131">Nome</span><span class="sxs-lookup"><span data-stu-id="55cc6-131">Name</span></span>| <span data-ttu-id="55cc6-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="55cc6-132">Type</span></span>| <span data-ttu-id="55cc6-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="55cc6-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="55cc6-134">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-134">String</span></span>|<span data-ttu-id="55cc6-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="55cc6-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="55cc6-136">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-136">String</span></span>|<span data-ttu-id="55cc6-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="55cc6-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55cc6-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55cc6-138">Requirements</span></span>

|<span data-ttu-id="55cc6-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="55cc6-139">Requirement</span></span>| <span data-ttu-id="55cc6-140">Valor</span><span class="sxs-lookup"><span data-stu-id="55cc6-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="55cc6-141">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55cc6-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55cc6-142">1.0</span><span class="sxs-lookup"><span data-stu-id="55cc6-142">1.0</span></span>|
|[<span data-ttu-id="55cc6-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55cc6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55cc6-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="55cc6-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="55cc6-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="55cc6-145">CoercionType :String</span></span>

<span data-ttu-id="55cc6-146">Especifica como forçar os dados retornados ou atribuídos pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="55cc6-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55cc6-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="55cc6-147">Type:</span></span>

*   <span data-ttu-id="55cc6-148">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55cc6-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55cc6-149">Properties:</span></span>

|<span data-ttu-id="55cc6-150">Nome</span><span class="sxs-lookup"><span data-stu-id="55cc6-150">Name</span></span>| <span data-ttu-id="55cc6-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="55cc6-151">Type</span></span>| <span data-ttu-id="55cc6-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="55cc6-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="55cc6-153">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-153">String</span></span>|<span data-ttu-id="55cc6-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="55cc6-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="55cc6-155">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-155">String</span></span>|<span data-ttu-id="55cc6-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="55cc6-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55cc6-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55cc6-157">Requirements</span></span>

|<span data-ttu-id="55cc6-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="55cc6-158">Requirement</span></span>| <span data-ttu-id="55cc6-159">Valor</span><span class="sxs-lookup"><span data-stu-id="55cc6-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="55cc6-160">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55cc6-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55cc6-161">1.0</span><span class="sxs-lookup"><span data-stu-id="55cc6-161">1.0</span></span>|
|[<span data-ttu-id="55cc6-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55cc6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55cc6-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="55cc6-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="55cc6-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="55cc6-164">EventType :String</span></span>

<span data-ttu-id="55cc6-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="55cc6-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="55cc6-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="55cc6-166">Type:</span></span>

*   <span data-ttu-id="55cc6-167">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55cc6-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55cc6-168">Properties:</span></span>

| <span data-ttu-id="55cc6-169">Nome</span><span class="sxs-lookup"><span data-stu-id="55cc6-169">Name</span></span> | <span data-ttu-id="55cc6-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="55cc6-170">Type</span></span> | <span data-ttu-id="55cc6-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="55cc6-171">Description</span></span> | <span data-ttu-id="55cc6-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="55cc6-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="55cc6-173">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-173">String</span></span> | <span data-ttu-id="55cc6-174">A data ou hora em que o compromisso ou série selecionada foi alterada.</span><span class="sxs-lookup"><span data-stu-id="55cc6-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="55cc6-175">1.7</span><span class="sxs-lookup"><span data-stu-id="55cc6-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="55cc6-176">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-176">String</span></span> | <span data-ttu-id="55cc6-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="55cc6-177">The selected item has changed.</span></span> | <span data-ttu-id="55cc6-178">1.5</span><span class="sxs-lookup"><span data-stu-id="55cc6-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="55cc6-179">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-179">String</span></span> | <span data-ttu-id="55cc6-180">A lista de destinatários do item ou local do compromisso selecionado foram alterados.</span><span class="sxs-lookup"><span data-stu-id="55cc6-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="55cc6-181">1.7</span><span class="sxs-lookup"><span data-stu-id="55cc6-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="55cc6-182">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-182">String</span></span> | <span data-ttu-id="55cc6-183">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="55cc6-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="55cc6-184">1.7</span><span class="sxs-lookup"><span data-stu-id="55cc6-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55cc6-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55cc6-185">Requirements</span></span>

|<span data-ttu-id="55cc6-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="55cc6-186">Requirement</span></span>| <span data-ttu-id="55cc6-187">Valor</span><span class="sxs-lookup"><span data-stu-id="55cc6-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="55cc6-188">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55cc6-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55cc6-189">1.5</span><span class="sxs-lookup"><span data-stu-id="55cc6-189">1.5</span></span> |
|[<span data-ttu-id="55cc6-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55cc6-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55cc6-191">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="55cc6-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="55cc6-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="55cc6-192">SourceProperty :String</span></span>

<span data-ttu-id="55cc6-193">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="55cc6-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55cc6-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="55cc6-194">Type:</span></span>

*   <span data-ttu-id="55cc6-195">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55cc6-196">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55cc6-196">Properties:</span></span>

|<span data-ttu-id="55cc6-197">Nome</span><span class="sxs-lookup"><span data-stu-id="55cc6-197">Name</span></span>| <span data-ttu-id="55cc6-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="55cc6-198">Type</span></span>| <span data-ttu-id="55cc6-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="55cc6-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="55cc6-200">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-200">String</span></span>|<span data-ttu-id="55cc6-201">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="55cc6-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="55cc6-202">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="55cc6-202">String</span></span>|<span data-ttu-id="55cc6-203">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="55cc6-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55cc6-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55cc6-204">Requirements</span></span>

|<span data-ttu-id="55cc6-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="55cc6-205">Requirement</span></span>| <span data-ttu-id="55cc6-206">Valor</span><span class="sxs-lookup"><span data-stu-id="55cc6-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="55cc6-207">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55cc6-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55cc6-208">1.0</span><span class="sxs-lookup"><span data-stu-id="55cc6-208">1.0</span></span>|
|[<span data-ttu-id="55cc6-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55cc6-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55cc6-210">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="55cc6-210">Compose or read</span></span>|