 

# <a name="office"></a><span data-ttu-id="f4646-101">Office</span><span class="sxs-lookup"><span data-stu-id="f4646-101">Office</span></span>

<span data-ttu-id="f4646-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f4646-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4646-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f4646-104">Requirements</span></span>

|<span data-ttu-id="f4646-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="f4646-105">Requirement</span></span>| <span data-ttu-id="f4646-106">Valor</span><span class="sxs-lookup"><span data-stu-id="f4646-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4646-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f4646-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4646-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f4646-108">1.0</span></span>|
|[<span data-ttu-id="f4646-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f4646-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f4646-110">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f4646-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f4646-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f4646-111">Namespaces</span></span>

<span data-ttu-id="f4646-112">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API dos suplementos do Office para uso na API do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f4646-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f4646-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f4646-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f4646-114">Membros</span><span class="sxs-lookup"><span data-stu-id="f4646-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f4646-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f4646-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="f4646-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f4646-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f4646-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f4646-117">Type:</span></span>

*   <span data-ttu-id="f4646-118">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4646-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f4646-119">Properties:</span></span>

|<span data-ttu-id="f4646-120">Nome</span><span class="sxs-lookup"><span data-stu-id="f4646-120">Name</span></span>| <span data-ttu-id="f4646-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="f4646-121">Type</span></span>| <span data-ttu-id="f4646-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="f4646-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f4646-123">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-123">String</span></span>|<span data-ttu-id="f4646-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="f4646-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f4646-125">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-125">String</span></span>|<span data-ttu-id="f4646-126">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="f4646-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4646-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f4646-127">Requirements</span></span>

|<span data-ttu-id="f4646-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="f4646-128">Requirement</span></span>| <span data-ttu-id="f4646-129">Valor</span><span class="sxs-lookup"><span data-stu-id="f4646-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4646-130">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f4646-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4646-131">1.0</span><span class="sxs-lookup"><span data-stu-id="f4646-131">1.0</span></span>|
|[<span data-ttu-id="f4646-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f4646-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f4646-133">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f4646-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="f4646-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f4646-134">CoercionType :String</span></span>

<span data-ttu-id="f4646-135">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="f4646-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f4646-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f4646-136">Type:</span></span>

*   <span data-ttu-id="f4646-137">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4646-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f4646-138">Properties:</span></span>

|<span data-ttu-id="f4646-139">Nome</span><span class="sxs-lookup"><span data-stu-id="f4646-139">Name</span></span>| <span data-ttu-id="f4646-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="f4646-140">Type</span></span>| <span data-ttu-id="f4646-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="f4646-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f4646-142">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-142">String</span></span>|<span data-ttu-id="f4646-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f4646-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f4646-144">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-144">String</span></span>|<span data-ttu-id="f4646-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f4646-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4646-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f4646-146">Requirements</span></span>

|<span data-ttu-id="f4646-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="f4646-147">Requirement</span></span>| <span data-ttu-id="f4646-148">Valor</span><span class="sxs-lookup"><span data-stu-id="f4646-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4646-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f4646-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4646-150">1.0</span><span class="sxs-lookup"><span data-stu-id="f4646-150">1.0</span></span>|
|[<span data-ttu-id="f4646-151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f4646-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f4646-152">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f4646-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="f4646-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f4646-153">SourceProperty :String</span></span>

<span data-ttu-id="f4646-154">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="f4646-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f4646-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f4646-155">Type:</span></span>

*   <span data-ttu-id="f4646-156">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4646-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f4646-157">Properties:</span></span>

|<span data-ttu-id="f4646-158">Nome</span><span class="sxs-lookup"><span data-stu-id="f4646-158">Name</span></span>| <span data-ttu-id="f4646-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="f4646-159">Type</span></span>| <span data-ttu-id="f4646-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="f4646-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f4646-161">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-161">String</span></span>|<span data-ttu-id="f4646-162">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f4646-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f4646-163">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f4646-163">String</span></span>|<span data-ttu-id="f4646-164">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f4646-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4646-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f4646-165">Requirements</span></span>

|<span data-ttu-id="f4646-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="f4646-166">Requirement</span></span>| <span data-ttu-id="f4646-167">Valor</span><span class="sxs-lookup"><span data-stu-id="f4646-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4646-168">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f4646-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4646-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f4646-169">1.0</span></span>|
|[<span data-ttu-id="f4646-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f4646-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f4646-171">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f4646-171">Compose or read</span></span>|