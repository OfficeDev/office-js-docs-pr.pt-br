 

# <a name="office"></a><span data-ttu-id="70088-101">Office</span><span class="sxs-lookup"><span data-stu-id="70088-101">Office</span></span>

<span data-ttu-id="70088-p101">O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="70088-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="70088-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70088-104">Requirements</span></span>

|<span data-ttu-id="70088-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="70088-105">Requirement</span></span>| <span data-ttu-id="70088-106">Valor</span><span class="sxs-lookup"><span data-stu-id="70088-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="70088-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70088-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70088-108">1.0</span><span class="sxs-lookup"><span data-stu-id="70088-108">1.0</span></span>|
|[<span data-ttu-id="70088-109">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="70088-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="70088-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="70088-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="70088-111">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="70088-111">Members and methods</span></span>

| <span data-ttu-id="70088-112">Membro</span><span class="sxs-lookup"><span data-stu-id="70088-112">Member</span></span> | <span data-ttu-id="70088-113">Tipo</span><span class="sxs-lookup"><span data-stu-id="70088-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="70088-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="70088-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="70088-115">Membro</span><span class="sxs-lookup"><span data-stu-id="70088-115">Member</span></span> |
| [<span data-ttu-id="70088-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="70088-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="70088-117">Membro</span><span class="sxs-lookup"><span data-stu-id="70088-117">Member</span></span> |
| [<span data-ttu-id="70088-118">EventType</span><span class="sxs-lookup"><span data-stu-id="70088-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="70088-119">Membro</span><span class="sxs-lookup"><span data-stu-id="70088-119">Member</span></span> |
| [<span data-ttu-id="70088-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="70088-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="70088-121">Membro</span><span class="sxs-lookup"><span data-stu-id="70088-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="70088-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="70088-122">Namespaces</span></span>

<span data-ttu-id="70088-123">[context](office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="70088-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="70088-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="70088-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="70088-125">Membros</span><span class="sxs-lookup"><span data-stu-id="70088-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="70088-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="70088-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="70088-127">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="70088-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="70088-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="70088-128">Type:</span></span>

*   <span data-ttu-id="70088-129">String</span><span class="sxs-lookup"><span data-stu-id="70088-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70088-130">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70088-130">Properties:</span></span>

|<span data-ttu-id="70088-131">Nome</span><span class="sxs-lookup"><span data-stu-id="70088-131">Name</span></span>| <span data-ttu-id="70088-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="70088-132">Type</span></span>| <span data-ttu-id="70088-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="70088-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="70088-134">String</span><span class="sxs-lookup"><span data-stu-id="70088-134">String</span></span>|<span data-ttu-id="70088-135">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="70088-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="70088-136">String</span><span class="sxs-lookup"><span data-stu-id="70088-136">String</span></span>|<span data-ttu-id="70088-137">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="70088-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70088-138">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70088-138">Requirements</span></span>

|<span data-ttu-id="70088-139">Requisito</span><span class="sxs-lookup"><span data-stu-id="70088-139">Requirement</span></span>| <span data-ttu-id="70088-140">Valor</span><span class="sxs-lookup"><span data-stu-id="70088-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="70088-141">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70088-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70088-142">1.0</span><span class="sxs-lookup"><span data-stu-id="70088-142">1.0</span></span>|
|[<span data-ttu-id="70088-143">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="70088-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="70088-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="70088-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="70088-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="70088-145">CoercionType :String</span></span>

<span data-ttu-id="70088-146">Especifica como forçar os dados retornados ou definir de acordo com o método invocado.</span><span class="sxs-lookup"><span data-stu-id="70088-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70088-147">Tipo:</span><span class="sxs-lookup"><span data-stu-id="70088-147">Type:</span></span>

*   <span data-ttu-id="70088-148">String</span><span class="sxs-lookup"><span data-stu-id="70088-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70088-149">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70088-149">Properties:</span></span>

|<span data-ttu-id="70088-150">Nome</span><span class="sxs-lookup"><span data-stu-id="70088-150">Name</span></span>| <span data-ttu-id="70088-151">Tipo</span><span class="sxs-lookup"><span data-stu-id="70088-151">Type</span></span>| <span data-ttu-id="70088-152">Descrição</span><span class="sxs-lookup"><span data-stu-id="70088-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="70088-153">String</span><span class="sxs-lookup"><span data-stu-id="70088-153">String</span></span>|<span data-ttu-id="70088-154">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="70088-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="70088-155">String</span><span class="sxs-lookup"><span data-stu-id="70088-155">String</span></span>|<span data-ttu-id="70088-156">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="70088-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70088-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70088-157">Requirements</span></span>

|<span data-ttu-id="70088-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="70088-158">Requirement</span></span>| <span data-ttu-id="70088-159">Valor</span><span class="sxs-lookup"><span data-stu-id="70088-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="70088-160">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70088-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70088-161">1.0</span><span class="sxs-lookup"><span data-stu-id="70088-161">1.0</span></span>|
|[<span data-ttu-id="70088-162">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="70088-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="70088-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="70088-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="70088-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="70088-164">EventType :String</span></span>

<span data-ttu-id="70088-165">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="70088-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="70088-166">Tipo:</span><span class="sxs-lookup"><span data-stu-id="70088-166">Type:</span></span>

*   <span data-ttu-id="70088-167">String</span><span class="sxs-lookup"><span data-stu-id="70088-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70088-168">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70088-168">Properties:</span></span>

| <span data-ttu-id="70088-169">Nome</span><span class="sxs-lookup"><span data-stu-id="70088-169">Name</span></span> | <span data-ttu-id="70088-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="70088-170">Type</span></span> | <span data-ttu-id="70088-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="70088-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="70088-172">String</span><span class="sxs-lookup"><span data-stu-id="70088-172">String</span></span> | <span data-ttu-id="70088-173">O item selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="70088-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="70088-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70088-174">Requirements</span></span>

|<span data-ttu-id="70088-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="70088-175">Requirement</span></span>| <span data-ttu-id="70088-176">Valor</span><span class="sxs-lookup"><span data-stu-id="70088-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="70088-177">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70088-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70088-178">1.5</span><span class="sxs-lookup"><span data-stu-id="70088-178">1.5</span></span> |
|[<span data-ttu-id="70088-179">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="70088-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="70088-180">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="70088-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="70088-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="70088-181">SourceProperty :String</span></span>

<span data-ttu-id="70088-182">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="70088-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70088-183">Tipo:</span><span class="sxs-lookup"><span data-stu-id="70088-183">Type:</span></span>

*   <span data-ttu-id="70088-184">String</span><span class="sxs-lookup"><span data-stu-id="70088-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70088-185">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70088-185">Properties:</span></span>

|<span data-ttu-id="70088-186">Nome</span><span class="sxs-lookup"><span data-stu-id="70088-186">Name</span></span>| <span data-ttu-id="70088-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="70088-187">Type</span></span>| <span data-ttu-id="70088-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="70088-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="70088-189">String</span><span class="sxs-lookup"><span data-stu-id="70088-189">String</span></span>|<span data-ttu-id="70088-190">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="70088-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="70088-191">String</span><span class="sxs-lookup"><span data-stu-id="70088-191">String</span></span>|<span data-ttu-id="70088-192">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="70088-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70088-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70088-193">Requirements</span></span>

|<span data-ttu-id="70088-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="70088-194">Requirement</span></span>| <span data-ttu-id="70088-195">Valor</span><span class="sxs-lookup"><span data-stu-id="70088-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="70088-196">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70088-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70088-197">1.0</span><span class="sxs-lookup"><span data-stu-id="70088-197">1.0</span></span>|
|[<span data-ttu-id="70088-198">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="70088-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="70088-199">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="70088-199">Compose or read</span></span>|