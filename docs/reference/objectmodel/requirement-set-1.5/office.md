# <a name="office"></a><span data-ttu-id="7262d-101">Office</span><span class="sxs-lookup"><span data-stu-id="7262d-101">Office</span></span>

<span data-ttu-id="7262d-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7262d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7262d-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7262d-104">Requirements</span></span>

|<span data-ttu-id="7262d-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="7262d-105">Requirement</span></span>| <span data-ttu-id="7262d-106">Valor</span><span class="sxs-lookup"><span data-stu-id="7262d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="7262d-107">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7262d-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7262d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="7262d-108">1.0</span></span>|
|[<span data-ttu-id="7262d-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7262d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7262d-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7262d-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7262d-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="7262d-111">Members and methods</span></span>

| <span data-ttu-id="7262d-112">Membro</span><span class="sxs-lookup"><span data-stu-id="7262d-112">Member</span></span> | <span data-ttu-id="7262d-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="7262d-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7262d-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7262d-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7262d-115">Membro</span><span class="sxs-lookup"><span data-stu-id="7262d-115">Member</span></span> |
| [<span data-ttu-id="7262d-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7262d-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7262d-117">Membro</span><span class="sxs-lookup"><span data-stu-id="7262d-117">Member</span></span> |
| [<span data-ttu-id="7262d-118">EventType</span><span class="sxs-lookup"><span data-stu-id="7262d-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7262d-119">Membro</span><span class="sxs-lookup"><span data-stu-id="7262d-119">Member</span></span> |
| [<span data-ttu-id="7262d-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7262d-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7262d-121">Membro</span><span class="sxs-lookup"><span data-stu-id="7262d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7262d-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7262d-122">Namespaces</span></span>

<span data-ttu-id="7262d-123">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API dos suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7262d-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7262d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="7262d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7262d-125">Membros</span><span class="sxs-lookup"><span data-stu-id="7262d-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="7262d-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="7262d-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="7262d-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="7262d-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7262d-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7262d-128">Type:</span></span>

*   <span data-ttu-id="7262d-129">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7262d-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7262d-130">Properties:</span></span>

|<span data-ttu-id="7262d-131">Nome</span><span class="sxs-lookup"><span data-stu-id="7262d-131">Name</span></span>| <span data-ttu-id="7262d-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="7262d-132">Type</span></span>| <span data-ttu-id="7262d-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="7262d-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7262d-134">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-134">String</span></span>|<span data-ttu-id="7262d-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="7262d-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7262d-136">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-136">String</span></span>|<span data-ttu-id="7262d-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="7262d-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7262d-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7262d-138">Requirements</span></span>

|<span data-ttu-id="7262d-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="7262d-139">Requirement</span></span>| <span data-ttu-id="7262d-140">Valor</span><span class="sxs-lookup"><span data-stu-id="7262d-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="7262d-141">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7262d-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7262d-142">1.0</span><span class="sxs-lookup"><span data-stu-id="7262d-142">1.0</span></span>|
|[<span data-ttu-id="7262d-143">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7262d-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7262d-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7262d-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="7262d-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="7262d-145">CoercionType :String</span></span>

<span data-ttu-id="7262d-146">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="7262d-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7262d-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7262d-147">Type:</span></span>

*   <span data-ttu-id="7262d-148">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7262d-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7262d-149">Properties:</span></span>

|<span data-ttu-id="7262d-150">Nome</span><span class="sxs-lookup"><span data-stu-id="7262d-150">Name</span></span>| <span data-ttu-id="7262d-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="7262d-151">Type</span></span>| <span data-ttu-id="7262d-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="7262d-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7262d-153">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-153">String</span></span>|<span data-ttu-id="7262d-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="7262d-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7262d-155">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-155">String</span></span>|<span data-ttu-id="7262d-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="7262d-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7262d-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7262d-157">Requirements</span></span>

|<span data-ttu-id="7262d-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="7262d-158">Requirement</span></span>| <span data-ttu-id="7262d-159">Valor</span><span class="sxs-lookup"><span data-stu-id="7262d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="7262d-160">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7262d-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7262d-161">1.0</span><span class="sxs-lookup"><span data-stu-id="7262d-161">1.0</span></span>|
|[<span data-ttu-id="7262d-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7262d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7262d-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7262d-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="7262d-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="7262d-164">EventType :String</span></span>

<span data-ttu-id="7262d-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="7262d-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7262d-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7262d-166">Type:</span></span>

*   <span data-ttu-id="7262d-167">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7262d-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7262d-168">Properties:</span></span>

| <span data-ttu-id="7262d-169">Nome</span><span class="sxs-lookup"><span data-stu-id="7262d-169">Name</span></span> | <span data-ttu-id="7262d-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="7262d-170">Type</span></span> | <span data-ttu-id="7262d-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="7262d-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="7262d-172">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-172">String</span></span> | <span data-ttu-id="7262d-173">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7262d-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7262d-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7262d-174">Requirements</span></span>

|<span data-ttu-id="7262d-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="7262d-175">Requirement</span></span>| <span data-ttu-id="7262d-176">Valor</span><span class="sxs-lookup"><span data-stu-id="7262d-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="7262d-177">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7262d-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7262d-178">1.5</span><span class="sxs-lookup"><span data-stu-id="7262d-178">1.5</span></span> |
|[<span data-ttu-id="7262d-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7262d-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7262d-180">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7262d-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="7262d-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="7262d-181">SourceProperty :String</span></span>

<span data-ttu-id="7262d-182">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="7262d-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7262d-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7262d-183">Type:</span></span>

*   <span data-ttu-id="7262d-184">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7262d-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7262d-185">Properties:</span></span>

|<span data-ttu-id="7262d-186">Nome</span><span class="sxs-lookup"><span data-stu-id="7262d-186">Name</span></span>| <span data-ttu-id="7262d-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="7262d-187">Type</span></span>| <span data-ttu-id="7262d-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="7262d-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7262d-189">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-189">String</span></span>|<span data-ttu-id="7262d-190">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7262d-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7262d-191">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7262d-191">String</span></span>|<span data-ttu-id="7262d-192">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7262d-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7262d-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7262d-193">Requirements</span></span>

|<span data-ttu-id="7262d-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="7262d-194">Requirement</span></span>| <span data-ttu-id="7262d-195">Valor</span><span class="sxs-lookup"><span data-stu-id="7262d-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="7262d-196">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7262d-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7262d-197">1.0</span><span class="sxs-lookup"><span data-stu-id="7262d-197">1.0</span></span>|
|[<span data-ttu-id="7262d-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7262d-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7262d-199">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7262d-199">Compose or read</span></span>|