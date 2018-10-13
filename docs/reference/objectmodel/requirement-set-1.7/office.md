 

# <a name="office"></a><span data-ttu-id="36bf0-101">Office</span><span class="sxs-lookup"><span data-stu-id="36bf0-101">Office</span></span>

<span data-ttu-id="36bf0-p101">O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="36bf0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="36bf0-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="36bf0-104">Requirements</span></span>

|<span data-ttu-id="36bf0-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="36bf0-105">Requirement</span></span>| <span data-ttu-id="36bf0-106">Valor</span><span class="sxs-lookup"><span data-stu-id="36bf0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="36bf0-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="36bf0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36bf0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="36bf0-108">1.0</span></span>|
|[<span data-ttu-id="36bf0-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="36bf0-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36bf0-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="36bf0-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="36bf0-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="36bf0-111">Members and methods</span></span>

| <span data-ttu-id="36bf0-112">Membro</span><span class="sxs-lookup"><span data-stu-id="36bf0-112">Member</span></span> | <span data-ttu-id="36bf0-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="36bf0-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="36bf0-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="36bf0-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="36bf0-115">Membro</span><span class="sxs-lookup"><span data-stu-id="36bf0-115">Member</span></span> |
| [<span data-ttu-id="36bf0-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="36bf0-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="36bf0-117">Membro</span><span class="sxs-lookup"><span data-stu-id="36bf0-117">Member</span></span> |
| [<span data-ttu-id="36bf0-118">EventType</span><span class="sxs-lookup"><span data-stu-id="36bf0-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="36bf0-119">Membro</span><span class="sxs-lookup"><span data-stu-id="36bf0-119">Member</span></span> |
| [<span data-ttu-id="36bf0-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="36bf0-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="36bf0-121">Membro</span><span class="sxs-lookup"><span data-stu-id="36bf0-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="36bf0-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="36bf0-122">Namespaces</span></span>

<span data-ttu-id="36bf0-123">[context](office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="36bf0-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="36bf0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="36bf0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="36bf0-125">Membros</span><span class="sxs-lookup"><span data-stu-id="36bf0-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="36bf0-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="36bf0-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="36bf0-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="36bf0-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="36bf0-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="36bf0-128">Type:</span></span>

*   <span data-ttu-id="36bf0-129">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36bf0-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="36bf0-130">Properties:</span></span>

|<span data-ttu-id="36bf0-131">Nome</span><span class="sxs-lookup"><span data-stu-id="36bf0-131">Name</span></span>| <span data-ttu-id="36bf0-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="36bf0-132">Type</span></span>| <span data-ttu-id="36bf0-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="36bf0-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="36bf0-134">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-134">String</span></span>|<span data-ttu-id="36bf0-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="36bf0-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="36bf0-136">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-136">String</span></span>|<span data-ttu-id="36bf0-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="36bf0-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36bf0-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="36bf0-138">Requirements</span></span>

|<span data-ttu-id="36bf0-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="36bf0-139">Requirement</span></span>| <span data-ttu-id="36bf0-140">Valor</span><span class="sxs-lookup"><span data-stu-id="36bf0-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="36bf0-141">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="36bf0-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36bf0-142">1.0</span><span class="sxs-lookup"><span data-stu-id="36bf0-142">1.0</span></span>|
|[<span data-ttu-id="36bf0-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="36bf0-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36bf0-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="36bf0-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="36bf0-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="36bf0-145">CoercionType :String</span></span>

<span data-ttu-id="36bf0-146">Especifica como forçar os dados retornados ou definidos de acordo com o método invocado.</span><span class="sxs-lookup"><span data-stu-id="36bf0-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="36bf0-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="36bf0-147">Type:</span></span>

*   <span data-ttu-id="36bf0-148">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36bf0-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="36bf0-149">Properties:</span></span>

|<span data-ttu-id="36bf0-150">Nome</span><span class="sxs-lookup"><span data-stu-id="36bf0-150">Name</span></span>| <span data-ttu-id="36bf0-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="36bf0-151">Type</span></span>| <span data-ttu-id="36bf0-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="36bf0-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="36bf0-153">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-153">String</span></span>|<span data-ttu-id="36bf0-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="36bf0-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="36bf0-155">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-155">String</span></span>|<span data-ttu-id="36bf0-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="36bf0-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36bf0-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="36bf0-157">Requirements</span></span>

|<span data-ttu-id="36bf0-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="36bf0-158">Requirement</span></span>| <span data-ttu-id="36bf0-159">Valor</span><span class="sxs-lookup"><span data-stu-id="36bf0-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="36bf0-160">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="36bf0-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36bf0-161">1.0</span><span class="sxs-lookup"><span data-stu-id="36bf0-161">1.0</span></span>|
|[<span data-ttu-id="36bf0-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="36bf0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36bf0-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="36bf0-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="36bf0-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="36bf0-164">EventType :String</span></span>

<span data-ttu-id="36bf0-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="36bf0-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="36bf0-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="36bf0-166">Type:</span></span>

*   <span data-ttu-id="36bf0-167">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36bf0-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="36bf0-168">Properties:</span></span>

| <span data-ttu-id="36bf0-169">Nome</span><span class="sxs-lookup"><span data-stu-id="36bf0-169">Name</span></span> | <span data-ttu-id="36bf0-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="36bf0-170">Type</span></span> | <span data-ttu-id="36bf0-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="36bf0-171">Description</span></span> | <span data-ttu-id="36bf0-172">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="36bf0-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="36bf0-173">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-173">String</span></span> | <span data-ttu-id="36bf0-174">A data ou hora do compromisso selecionado ou série foi alterada.</span><span class="sxs-lookup"><span data-stu-id="36bf0-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="36bf0-175">1.7</span><span class="sxs-lookup"><span data-stu-id="36bf0-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="36bf0-176">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-176">String</span></span> | <span data-ttu-id="36bf0-177">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="36bf0-177">The selected item has changed.</span></span> | <span data-ttu-id="36bf0-178">1.5</span><span class="sxs-lookup"><span data-stu-id="36bf0-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="36bf0-179">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-179">String</span></span> | <span data-ttu-id="36bf0-180">A lista de destinatários do item ou local do compromisso selecionado foram alterados.</span><span class="sxs-lookup"><span data-stu-id="36bf0-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="36bf0-181">1.7</span><span class="sxs-lookup"><span data-stu-id="36bf0-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="36bf0-182">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-182">String</span></span> | <span data-ttu-id="36bf0-183">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="36bf0-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="36bf0-184">1.7</span><span class="sxs-lookup"><span data-stu-id="36bf0-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="36bf0-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="36bf0-185">Requirements</span></span>

|<span data-ttu-id="36bf0-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="36bf0-186">Requirement</span></span>| <span data-ttu-id="36bf0-187">Valor</span><span class="sxs-lookup"><span data-stu-id="36bf0-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="36bf0-188">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="36bf0-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36bf0-189">1.5</span><span class="sxs-lookup"><span data-stu-id="36bf0-189">1.5</span></span> |
|[<span data-ttu-id="36bf0-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="36bf0-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36bf0-191">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="36bf0-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="36bf0-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="36bf0-192">SourceProperty :String</span></span>

<span data-ttu-id="36bf0-193">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="36bf0-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="36bf0-194">Tipo:</span><span class="sxs-lookup"><span data-stu-id="36bf0-194">Type:</span></span>

*   <span data-ttu-id="36bf0-195">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36bf0-196">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="36bf0-196">Properties:</span></span>

|<span data-ttu-id="36bf0-197">Nome</span><span class="sxs-lookup"><span data-stu-id="36bf0-197">Name</span></span>| <span data-ttu-id="36bf0-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="36bf0-198">Type</span></span>| <span data-ttu-id="36bf0-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="36bf0-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="36bf0-200">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-200">String</span></span>|<span data-ttu-id="36bf0-201">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="36bf0-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="36bf0-202">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="36bf0-202">String</span></span>|<span data-ttu-id="36bf0-203">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="36bf0-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36bf0-204">Requisitos</span><span class="sxs-lookup"><span data-stu-id="36bf0-204">Requirements</span></span>

|<span data-ttu-id="36bf0-205">Requisito</span><span class="sxs-lookup"><span data-stu-id="36bf0-205">Requirement</span></span>| <span data-ttu-id="36bf0-206">Valor</span><span class="sxs-lookup"><span data-stu-id="36bf0-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="36bf0-207">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="36bf0-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36bf0-208">1.0</span><span class="sxs-lookup"><span data-stu-id="36bf0-208">1.0</span></span>|
|[<span data-ttu-id="36bf0-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="36bf0-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36bf0-210">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="36bf0-210">Compose or read</span></span>|