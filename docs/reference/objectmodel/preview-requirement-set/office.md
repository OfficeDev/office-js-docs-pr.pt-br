 

# <a name="office"></a><span data-ttu-id="f9680-101">Office</span><span class="sxs-lookup"><span data-stu-id="f9680-101">Office</span></span>

<span data-ttu-id="f9680-p101">O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f9680-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9680-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9680-104">Requirements</span></span>

|<span data-ttu-id="f9680-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9680-105">Requirement</span></span>| <span data-ttu-id="f9680-106">Valor</span><span class="sxs-lookup"><span data-stu-id="f9680-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9680-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9680-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9680-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f9680-108">1.0</span></span>|
|[<span data-ttu-id="f9680-109">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f9680-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9680-110">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9680-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f9680-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="f9680-111">Members and methods</span></span>

| <span data-ttu-id="f9680-112">Membro</span><span class="sxs-lookup"><span data-stu-id="f9680-112">Member</span></span> | <span data-ttu-id="f9680-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9680-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f9680-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f9680-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f9680-115">Membro</span><span class="sxs-lookup"><span data-stu-id="f9680-115">Member</span></span> |
| [<span data-ttu-id="f9680-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9680-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f9680-117">Membro</span><span class="sxs-lookup"><span data-stu-id="f9680-117">Member</span></span> |
| [<span data-ttu-id="f9680-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f9680-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f9680-119">Membro</span><span class="sxs-lookup"><span data-stu-id="f9680-119">Member</span></span> |
| [<span data-ttu-id="f9680-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f9680-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f9680-121">Membro</span><span class="sxs-lookup"><span data-stu-id="f9680-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f9680-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f9680-122">Namespaces</span></span>

<span data-ttu-id="f9680-123">[context](office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f9680-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f9680-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f9680-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f9680-125">Membros</span><span class="sxs-lookup"><span data-stu-id="f9680-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f9680-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f9680-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f9680-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f9680-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f9680-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9680-128">Type:</span></span>

*   <span data-ttu-id="f9680-129">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9680-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9680-130">Properties:</span></span>

|<span data-ttu-id="f9680-131">Nome</span><span class="sxs-lookup"><span data-stu-id="f9680-131">Name</span></span>| <span data-ttu-id="f9680-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9680-132">Type</span></span>| <span data-ttu-id="f9680-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9680-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f9680-134">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-134">String</span></span>|<span data-ttu-id="f9680-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="f9680-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f9680-136">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-136">String</span></span>|<span data-ttu-id="f9680-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="f9680-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9680-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9680-138">Requirements</span></span>

|<span data-ttu-id="f9680-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9680-139">Requirement</span></span>| <span data-ttu-id="f9680-140">Valor</span><span class="sxs-lookup"><span data-stu-id="f9680-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9680-141">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9680-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9680-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f9680-142">1.0</span></span>|
|[<span data-ttu-id="f9680-143">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f9680-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9680-144">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9680-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f9680-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f9680-145">CoercionType :String</span></span>

<span data-ttu-id="f9680-146">Especifica como forçar os dados retornados ou definidos de acordo com o método invocado.</span><span class="sxs-lookup"><span data-stu-id="f9680-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9680-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9680-147">Type:</span></span>

*   <span data-ttu-id="f9680-148">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9680-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9680-149">Properties:</span></span>

|<span data-ttu-id="f9680-150">Nome</span><span class="sxs-lookup"><span data-stu-id="f9680-150">Name</span></span>| <span data-ttu-id="f9680-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9680-151">Type</span></span>| <span data-ttu-id="f9680-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9680-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f9680-153">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-153">String</span></span>|<span data-ttu-id="f9680-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f9680-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f9680-155">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-155">String</span></span>|<span data-ttu-id="f9680-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f9680-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9680-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9680-157">Requirements</span></span>

|<span data-ttu-id="f9680-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9680-158">Requirement</span></span>| <span data-ttu-id="f9680-159">Valor</span><span class="sxs-lookup"><span data-stu-id="f9680-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9680-160">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9680-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9680-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f9680-161">1.0</span></span>|
|[<span data-ttu-id="f9680-162">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f9680-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9680-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9680-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f9680-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f9680-164">EventType :String</span></span>

<span data-ttu-id="f9680-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="f9680-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f9680-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9680-166">Type:</span></span>

*   <span data-ttu-id="f9680-167">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9680-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9680-168">Properties:</span></span>

| <span data-ttu-id="f9680-169">Nome</span><span class="sxs-lookup"><span data-stu-id="f9680-169">Name</span></span> | <span data-ttu-id="f9680-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9680-170">Type</span></span> | <span data-ttu-id="f9680-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9680-171">Description</span></span> | <span data-ttu-id="f9680-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="f9680-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f9680-173">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-173">String</span></span> | <span data-ttu-id="f9680-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f9680-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f9680-175">1.7</span><span class="sxs-lookup"><span data-stu-id="f9680-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="f9680-176">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-176">String</span></span> | <span data-ttu-id="f9680-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f9680-177">The selected item has changed.</span></span> | <span data-ttu-id="f9680-178">1.5</span><span class="sxs-lookup"><span data-stu-id="f9680-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f9680-179">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-179">String</span></span> | <span data-ttu-id="f9680-180">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f9680-180">The selected item has changed.</span></span> | <span data-ttu-id="f9680-181">Visualização</span><span class="sxs-lookup"><span data-stu-id="f9680-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f9680-182">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-182">String</span></span> | <span data-ttu-id="f9680-183">A lista de destinatários do item ou local do compromisso selecionado foram alterados.</span><span class="sxs-lookup"><span data-stu-id="f9680-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f9680-184">1.7</span><span class="sxs-lookup"><span data-stu-id="f9680-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f9680-185">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-185">String</span></span> | <span data-ttu-id="f9680-186">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f9680-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f9680-187">1.7</span><span class="sxs-lookup"><span data-stu-id="f9680-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9680-188">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9680-188">Requirements</span></span>

|<span data-ttu-id="f9680-189">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9680-189">Requirement</span></span>| <span data-ttu-id="f9680-190">Valor</span><span class="sxs-lookup"><span data-stu-id="f9680-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9680-191">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9680-191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9680-192">1.5</span><span class="sxs-lookup"><span data-stu-id="f9680-192">1.5</span></span> |
|[<span data-ttu-id="f9680-193">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f9680-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9680-194">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9680-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f9680-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f9680-195">SourceProperty :String</span></span>

<span data-ttu-id="f9680-196">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="f9680-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9680-197">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9680-197">Type:</span></span>

*   <span data-ttu-id="f9680-198">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9680-199">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9680-199">Properties:</span></span>

|<span data-ttu-id="f9680-200">Nome</span><span class="sxs-lookup"><span data-stu-id="f9680-200">Name</span></span>| <span data-ttu-id="f9680-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9680-201">Type</span></span>| <span data-ttu-id="f9680-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9680-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f9680-203">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-203">String</span></span>|<span data-ttu-id="f9680-204">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9680-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f9680-205">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9680-205">String</span></span>|<span data-ttu-id="f9680-206">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9680-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9680-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9680-207">Requirements</span></span>

|<span data-ttu-id="f9680-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9680-208">Requirement</span></span>| <span data-ttu-id="f9680-209">Valor</span><span class="sxs-lookup"><span data-stu-id="f9680-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9680-210">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9680-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9680-211">1.0</span><span class="sxs-lookup"><span data-stu-id="f9680-211">1.0</span></span>|
|[<span data-ttu-id="f9680-212">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f9680-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9680-213">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9680-213">Compose or read</span></span>|