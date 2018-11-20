# <a name="office"></a><span data-ttu-id="f9065-101">Office</span><span class="sxs-lookup"><span data-stu-id="f9065-101">Office</span></span>

<span data-ttu-id="f9065-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f9065-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f9065-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9065-104">Requirements</span></span>

|<span data-ttu-id="f9065-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9065-105">Requirement</span></span>| <span data-ttu-id="f9065-106">Valor</span><span class="sxs-lookup"><span data-stu-id="f9065-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9065-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9065-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9065-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f9065-108">1.0</span></span>|
|[<span data-ttu-id="f9065-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9065-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9065-110">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9065-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f9065-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="f9065-111">Members and methods</span></span>

| <span data-ttu-id="f9065-112">Membro</span><span class="sxs-lookup"><span data-stu-id="f9065-112">Member</span></span> | <span data-ttu-id="f9065-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9065-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f9065-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f9065-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f9065-115">Membro</span><span class="sxs-lookup"><span data-stu-id="f9065-115">Member</span></span> |
| [<span data-ttu-id="f9065-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f9065-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f9065-117">Membro</span><span class="sxs-lookup"><span data-stu-id="f9065-117">Member</span></span> |
| [<span data-ttu-id="f9065-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f9065-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f9065-119">Membro</span><span class="sxs-lookup"><span data-stu-id="f9065-119">Member</span></span> |
| [<span data-ttu-id="f9065-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f9065-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f9065-121">Membro</span><span class="sxs-lookup"><span data-stu-id="f9065-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f9065-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f9065-122">Namespaces</span></span>

<span data-ttu-id="f9065-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f9065-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f9065-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f9065-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f9065-125">Membros</span><span class="sxs-lookup"><span data-stu-id="f9065-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f9065-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f9065-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f9065-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f9065-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f9065-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9065-128">Type:</span></span>

*   <span data-ttu-id="f9065-129">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9065-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9065-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9065-130">Properties:</span></span>

|<span data-ttu-id="f9065-131">Nome</span><span class="sxs-lookup"><span data-stu-id="f9065-131">Name</span></span>| <span data-ttu-id="f9065-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9065-132">Type</span></span>| <span data-ttu-id="f9065-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9065-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f9065-134">String</span><span class="sxs-lookup"><span data-stu-id="f9065-134">String</span></span>|<span data-ttu-id="f9065-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="f9065-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f9065-136">String</span><span class="sxs-lookup"><span data-stu-id="f9065-136">String</span></span>|<span data-ttu-id="f9065-137">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="f9065-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9065-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9065-138">Requirements</span></span>

|<span data-ttu-id="f9065-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9065-139">Requirement</span></span>| <span data-ttu-id="f9065-140">Valor</span><span class="sxs-lookup"><span data-stu-id="f9065-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9065-141">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9065-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9065-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f9065-142">1.0</span></span>|
|[<span data-ttu-id="f9065-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9065-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9065-144">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9065-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f9065-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f9065-145">CoercionType :String</span></span>

<span data-ttu-id="f9065-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="f9065-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9065-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9065-147">Type:</span></span>

*   <span data-ttu-id="f9065-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9065-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9065-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9065-149">Properties:</span></span>

|<span data-ttu-id="f9065-150">Nome</span><span class="sxs-lookup"><span data-stu-id="f9065-150">Name</span></span>| <span data-ttu-id="f9065-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9065-151">Type</span></span>| <span data-ttu-id="f9065-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9065-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f9065-153">String</span><span class="sxs-lookup"><span data-stu-id="f9065-153">String</span></span>|<span data-ttu-id="f9065-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f9065-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f9065-155">String</span><span class="sxs-lookup"><span data-stu-id="f9065-155">String</span></span>|<span data-ttu-id="f9065-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f9065-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9065-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9065-157">Requirements</span></span>

|<span data-ttu-id="f9065-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9065-158">Requirement</span></span>| <span data-ttu-id="f9065-159">Valor</span><span class="sxs-lookup"><span data-stu-id="f9065-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9065-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9065-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9065-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f9065-161">1.0</span></span>|
|[<span data-ttu-id="f9065-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9065-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9065-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9065-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f9065-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f9065-164">EventType :String</span></span>

<span data-ttu-id="f9065-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="f9065-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f9065-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9065-166">Type:</span></span>

*   <span data-ttu-id="f9065-167">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9065-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9065-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9065-168">Properties:</span></span>

| <span data-ttu-id="f9065-169">Nome</span><span class="sxs-lookup"><span data-stu-id="f9065-169">Name</span></span> | <span data-ttu-id="f9065-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9065-170">Type</span></span> | <span data-ttu-id="f9065-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9065-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="f9065-172">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9065-172">String</span></span> | <span data-ttu-id="f9065-173">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="f9065-173">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f9065-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9065-174">Requirements</span></span>

|<span data-ttu-id="f9065-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9065-175">Requirement</span></span>| <span data-ttu-id="f9065-176">Valor</span><span class="sxs-lookup"><span data-stu-id="f9065-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9065-177">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9065-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9065-178">1.5</span><span class="sxs-lookup"><span data-stu-id="f9065-178">1.5</span></span> |
|[<span data-ttu-id="f9065-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9065-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9065-180">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9065-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f9065-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f9065-181">SourceProperty :String</span></span>

<span data-ttu-id="f9065-182">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="f9065-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f9065-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f9065-183">Type:</span></span>

*   <span data-ttu-id="f9065-184">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f9065-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f9065-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f9065-185">Properties:</span></span>

|<span data-ttu-id="f9065-186">Nome</span><span class="sxs-lookup"><span data-stu-id="f9065-186">Name</span></span>| <span data-ttu-id="f9065-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="f9065-187">Type</span></span>| <span data-ttu-id="f9065-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="f9065-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f9065-189">String</span><span class="sxs-lookup"><span data-stu-id="f9065-189">String</span></span>|<span data-ttu-id="f9065-190">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9065-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f9065-191">String</span><span class="sxs-lookup"><span data-stu-id="f9065-191">String</span></span>|<span data-ttu-id="f9065-192">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f9065-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f9065-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f9065-193">Requirements</span></span>

|<span data-ttu-id="f9065-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="f9065-194">Requirement</span></span>| <span data-ttu-id="f9065-195">Valor</span><span class="sxs-lookup"><span data-stu-id="f9065-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="f9065-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f9065-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f9065-197">1.0</span><span class="sxs-lookup"><span data-stu-id="f9065-197">1.0</span></span>|
|[<span data-ttu-id="f9065-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f9065-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f9065-199">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="f9065-199">Compose or read</span></span>|