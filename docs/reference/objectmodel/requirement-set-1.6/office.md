 

# <a name="office"></a><span data-ttu-id="dc411-101">Office</span><span class="sxs-lookup"><span data-stu-id="dc411-101">Office</span></span>

<span data-ttu-id="dc411-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="dc411-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc411-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc411-104">Requirements</span></span>

|<span data-ttu-id="dc411-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc411-105">Requirement</span></span>| <span data-ttu-id="dc411-106">Valor</span><span class="sxs-lookup"><span data-stu-id="dc411-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc411-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc411-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc411-108">1.0</span><span class="sxs-lookup"><span data-stu-id="dc411-108">1.0</span></span>|
|[<span data-ttu-id="dc411-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc411-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc411-110">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc411-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dc411-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="dc411-111">Members and methods</span></span>

| <span data-ttu-id="dc411-112">Membro</span><span class="sxs-lookup"><span data-stu-id="dc411-112">Member</span></span> | <span data-ttu-id="dc411-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="dc411-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dc411-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dc411-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dc411-115">Membro</span><span class="sxs-lookup"><span data-stu-id="dc411-115">Member</span></span> |
| [<span data-ttu-id="dc411-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dc411-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dc411-117">Membro</span><span class="sxs-lookup"><span data-stu-id="dc411-117">Member</span></span> |
| [<span data-ttu-id="dc411-118">EventType</span><span class="sxs-lookup"><span data-stu-id="dc411-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dc411-119">Membro</span><span class="sxs-lookup"><span data-stu-id="dc411-119">Member</span></span> |
| [<span data-ttu-id="dc411-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dc411-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dc411-121">Membro</span><span class="sxs-lookup"><span data-stu-id="dc411-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="dc411-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="dc411-122">Namespaces</span></span>

<span data-ttu-id="dc411-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="dc411-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="dc411-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="dc411-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="dc411-125">Membros</span><span class="sxs-lookup"><span data-stu-id="dc411-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="dc411-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="dc411-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="dc411-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="dc411-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dc411-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc411-128">Type:</span></span>

*   <span data-ttu-id="dc411-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc411-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc411-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="dc411-130">Properties:</span></span>

|<span data-ttu-id="dc411-131">Nome</span><span class="sxs-lookup"><span data-stu-id="dc411-131">Name</span></span>| <span data-ttu-id="dc411-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="dc411-132">Type</span></span>| <span data-ttu-id="dc411-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc411-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dc411-134">String</span><span class="sxs-lookup"><span data-stu-id="dc411-134">String</span></span>|<span data-ttu-id="dc411-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="dc411-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dc411-136">String</span><span class="sxs-lookup"><span data-stu-id="dc411-136">String</span></span>|<span data-ttu-id="dc411-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="dc411-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc411-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc411-138">Requirements</span></span>

|<span data-ttu-id="dc411-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc411-139">Requirement</span></span>| <span data-ttu-id="dc411-140">Valor</span><span class="sxs-lookup"><span data-stu-id="dc411-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc411-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc411-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc411-142">1.0</span><span class="sxs-lookup"><span data-stu-id="dc411-142">1.0</span></span>|
|[<span data-ttu-id="dc411-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc411-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc411-144">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc411-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="dc411-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="dc411-145">CoercionType :String</span></span>

<span data-ttu-id="dc411-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="dc411-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc411-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc411-147">Type:</span></span>

*   <span data-ttu-id="dc411-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc411-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc411-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="dc411-149">Properties:</span></span>

|<span data-ttu-id="dc411-150">Nome</span><span class="sxs-lookup"><span data-stu-id="dc411-150">Name</span></span>| <span data-ttu-id="dc411-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="dc411-151">Type</span></span>| <span data-ttu-id="dc411-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc411-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dc411-153">String</span><span class="sxs-lookup"><span data-stu-id="dc411-153">String</span></span>|<span data-ttu-id="dc411-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="dc411-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dc411-155">String</span><span class="sxs-lookup"><span data-stu-id="dc411-155">String</span></span>|<span data-ttu-id="dc411-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="dc411-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc411-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc411-157">Requirements</span></span>

|<span data-ttu-id="dc411-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc411-158">Requirement</span></span>| <span data-ttu-id="dc411-159">Valor</span><span class="sxs-lookup"><span data-stu-id="dc411-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc411-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc411-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc411-161">1.0</span><span class="sxs-lookup"><span data-stu-id="dc411-161">1.0</span></span>|
|[<span data-ttu-id="dc411-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc411-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc411-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc411-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="dc411-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="dc411-164">EventType :String</span></span>

<span data-ttu-id="dc411-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="dc411-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dc411-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc411-166">Type:</span></span>

*   <span data-ttu-id="dc411-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc411-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc411-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="dc411-168">Properties:</span></span>

| <span data-ttu-id="dc411-169">Nome</span><span class="sxs-lookup"><span data-stu-id="dc411-169">Name</span></span> | <span data-ttu-id="dc411-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="dc411-170">Type</span></span> | <span data-ttu-id="dc411-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc411-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="dc411-172">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc411-172">String</span></span> | <span data-ttu-id="dc411-173">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="dc411-173">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc411-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc411-174">Requirements</span></span>

|<span data-ttu-id="dc411-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc411-175">Requirement</span></span>| <span data-ttu-id="dc411-176">Valor</span><span class="sxs-lookup"><span data-stu-id="dc411-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc411-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc411-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc411-178">1.5</span><span class="sxs-lookup"><span data-stu-id="dc411-178">1.5</span></span> |
|[<span data-ttu-id="dc411-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc411-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc411-180">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc411-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="dc411-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="dc411-181">SourceProperty :String</span></span>

<span data-ttu-id="dc411-182">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="dc411-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc411-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc411-183">Type:</span></span>

*   <span data-ttu-id="dc411-184">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc411-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc411-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="dc411-185">Properties:</span></span>

|<span data-ttu-id="dc411-186">Nome</span><span class="sxs-lookup"><span data-stu-id="dc411-186">Name</span></span>| <span data-ttu-id="dc411-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="dc411-187">Type</span></span>| <span data-ttu-id="dc411-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc411-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dc411-189">String</span><span class="sxs-lookup"><span data-stu-id="dc411-189">String</span></span>|<span data-ttu-id="dc411-190">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="dc411-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dc411-191">String</span><span class="sxs-lookup"><span data-stu-id="dc411-191">String</span></span>|<span data-ttu-id="dc411-192">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="dc411-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc411-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc411-193">Requirements</span></span>

|<span data-ttu-id="dc411-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc411-194">Requirement</span></span>| <span data-ttu-id="dc411-195">Valor</span><span class="sxs-lookup"><span data-stu-id="dc411-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc411-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc411-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc411-197">1.0</span><span class="sxs-lookup"><span data-stu-id="dc411-197">1.0</span></span>|
|[<span data-ttu-id="dc411-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc411-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc411-199">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc411-199">Compose or read</span></span>|