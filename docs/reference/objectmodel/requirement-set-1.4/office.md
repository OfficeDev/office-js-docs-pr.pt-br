 

# <a name="office"></a><span data-ttu-id="75b42-101">Office</span><span class="sxs-lookup"><span data-stu-id="75b42-101">Office</span></span>

<span data-ttu-id="75b42-p101">O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="75b42-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="75b42-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="75b42-104">Requirements</span></span>

|<span data-ttu-id="75b42-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="75b42-105">Requirement</span></span>| <span data-ttu-id="75b42-106">Valor</span><span class="sxs-lookup"><span data-stu-id="75b42-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="75b42-107">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="75b42-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75b42-108">1.0</span><span class="sxs-lookup"><span data-stu-id="75b42-108">1.0</span></span>|
|[<span data-ttu-id="75b42-109">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="75b42-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75b42-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="75b42-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="75b42-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="75b42-111">Namespaces</span></span>

<span data-ttu-id="75b42-112">[context](Office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="75b42-112">[context](Office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="75b42-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="75b42-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="75b42-114">Membros</span><span class="sxs-lookup"><span data-stu-id="75b42-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="75b42-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="75b42-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="75b42-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="75b42-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="75b42-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="75b42-117">Type:</span></span>

*   <span data-ttu-id="75b42-118">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="75b42-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="75b42-119">Properties:</span></span>

|<span data-ttu-id="75b42-120">Nome</span><span class="sxs-lookup"><span data-stu-id="75b42-120">Name</span></span>| <span data-ttu-id="75b42-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="75b42-121">Type</span></span>| <span data-ttu-id="75b42-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="75b42-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="75b42-123">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-123">String</span></span>|<span data-ttu-id="75b42-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="75b42-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="75b42-125">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-125">String</span></span>|<span data-ttu-id="75b42-126">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="75b42-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75b42-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="75b42-127">Requirements</span></span>

|<span data-ttu-id="75b42-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="75b42-128">Requirement</span></span>| <span data-ttu-id="75b42-129">Valor</span><span class="sxs-lookup"><span data-stu-id="75b42-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="75b42-130">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="75b42-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75b42-131">1.0</span><span class="sxs-lookup"><span data-stu-id="75b42-131">1.0</span></span>|
|[<span data-ttu-id="75b42-132">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="75b42-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75b42-133">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="75b42-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="75b42-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="75b42-134">CoercionType :String</span></span>

<span data-ttu-id="75b42-135">Especifica como forçar os dados retornados ou atribuídos pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="75b42-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="75b42-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="75b42-136">Type:</span></span>

*   <span data-ttu-id="75b42-137">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="75b42-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="75b42-138">Properties:</span></span>

|<span data-ttu-id="75b42-139">Nome</span><span class="sxs-lookup"><span data-stu-id="75b42-139">Name</span></span>| <span data-ttu-id="75b42-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="75b42-140">Type</span></span>| <span data-ttu-id="75b42-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="75b42-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="75b42-142">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-142">String</span></span>|<span data-ttu-id="75b42-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="75b42-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="75b42-144">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-144">String</span></span>|<span data-ttu-id="75b42-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="75b42-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75b42-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="75b42-146">Requirements</span></span>

|<span data-ttu-id="75b42-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="75b42-147">Requirement</span></span>| <span data-ttu-id="75b42-148">Valor</span><span class="sxs-lookup"><span data-stu-id="75b42-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="75b42-149">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="75b42-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75b42-150">1.0</span><span class="sxs-lookup"><span data-stu-id="75b42-150">1.0</span></span>|
|[<span data-ttu-id="75b42-151">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="75b42-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75b42-152">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="75b42-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="75b42-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="75b42-153">SourceProperty :String</span></span>

<span data-ttu-id="75b42-154">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="75b42-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="75b42-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="75b42-155">Type:</span></span>

*   <span data-ttu-id="75b42-156">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="75b42-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="75b42-157">Properties:</span></span>

|<span data-ttu-id="75b42-158">Nome</span><span class="sxs-lookup"><span data-stu-id="75b42-158">Name</span></span>| <span data-ttu-id="75b42-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="75b42-159">Type</span></span>| <span data-ttu-id="75b42-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="75b42-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="75b42-161">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-161">String</span></span>|<span data-ttu-id="75b42-162">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="75b42-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="75b42-163">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="75b42-163">String</span></span>|<span data-ttu-id="75b42-164">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="75b42-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75b42-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="75b42-165">Requirements</span></span>

|<span data-ttu-id="75b42-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="75b42-166">Requirement</span></span>| <span data-ttu-id="75b42-167">Valor</span><span class="sxs-lookup"><span data-stu-id="75b42-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="75b42-168">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="75b42-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75b42-169">1.0</span><span class="sxs-lookup"><span data-stu-id="75b42-169">1.0</span></span>|
|[<span data-ttu-id="75b42-170">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="75b42-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75b42-171">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="75b42-171">Compose or read</span></span>|