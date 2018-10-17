 

# <a name="office"></a><span data-ttu-id="171d3-101">Office</span><span class="sxs-lookup"><span data-stu-id="171d3-101">Office</span></span>

<span data-ttu-id="171d3-p101">O namespace Office fornece interfaces compartilhadas que são usadas por suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="171d3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="171d3-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="171d3-104">Requirements</span></span>

|<span data-ttu-id="171d3-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="171d3-105">Requirement</span></span>| <span data-ttu-id="171d3-106">Valor</span><span class="sxs-lookup"><span data-stu-id="171d3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="171d3-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="171d3-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="171d3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="171d3-108">1.0</span></span>|
|[<span data-ttu-id="171d3-109">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="171d3-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="171d3-110">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="171d3-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="171d3-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="171d3-111">Namespaces</span></span>

<span data-ttu-id="171d3-112">[context](Office.context.md): fornece interfaces compartilhadas do namespace do contexto da API de suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="171d3-112">[context](Office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="171d3-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="171d3-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="171d3-114">Membros</span><span class="sxs-lookup"><span data-stu-id="171d3-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="171d3-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="171d3-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="171d3-116">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="171d3-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="171d3-117">Tipo:</span><span class="sxs-lookup"><span data-stu-id="171d3-117">Type:</span></span>

*   <span data-ttu-id="171d3-118">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="171d3-119">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="171d3-119">Properties:</span></span>

|<span data-ttu-id="171d3-120">Nome</span><span class="sxs-lookup"><span data-stu-id="171d3-120">Name</span></span>| <span data-ttu-id="171d3-121">Tipo</span><span class="sxs-lookup"><span data-stu-id="171d3-121">Type</span></span>| <span data-ttu-id="171d3-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="171d3-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="171d3-123">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-123">String</span></span>|<span data-ttu-id="171d3-124">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="171d3-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="171d3-125">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-125">String</span></span>|<span data-ttu-id="171d3-126">A chamada falhou.</span><span class="sxs-lookup"><span data-stu-id="171d3-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="171d3-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="171d3-127">Requirements</span></span>

|<span data-ttu-id="171d3-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="171d3-128">Requirement</span></span>| <span data-ttu-id="171d3-129">Valor</span><span class="sxs-lookup"><span data-stu-id="171d3-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="171d3-130">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="171d3-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="171d3-131">1.0</span><span class="sxs-lookup"><span data-stu-id="171d3-131">1.0</span></span>|
|[<span data-ttu-id="171d3-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="171d3-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="171d3-133">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="171d3-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="171d3-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="171d3-134">CoercionType :String</span></span>

<span data-ttu-id="171d3-135">Especifica como forçar os dados retornados ou definidos de acordo com o método invocado.</span><span class="sxs-lookup"><span data-stu-id="171d3-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="171d3-136">Tipo:</span><span class="sxs-lookup"><span data-stu-id="171d3-136">Type:</span></span>

*   <span data-ttu-id="171d3-137">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="171d3-138">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="171d3-138">Properties:</span></span>

|<span data-ttu-id="171d3-139">Nome</span><span class="sxs-lookup"><span data-stu-id="171d3-139">Name</span></span>| <span data-ttu-id="171d3-140">Tipo</span><span class="sxs-lookup"><span data-stu-id="171d3-140">Type</span></span>| <span data-ttu-id="171d3-141">Descrição</span><span class="sxs-lookup"><span data-stu-id="171d3-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="171d3-142">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-142">String</span></span>|<span data-ttu-id="171d3-143">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="171d3-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="171d3-144">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-144">String</span></span>|<span data-ttu-id="171d3-145">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="171d3-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="171d3-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="171d3-146">Requirements</span></span>

|<span data-ttu-id="171d3-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="171d3-147">Requirement</span></span>| <span data-ttu-id="171d3-148">Valor</span><span class="sxs-lookup"><span data-stu-id="171d3-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="171d3-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="171d3-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="171d3-150">1.0</span><span class="sxs-lookup"><span data-stu-id="171d3-150">1.0</span></span>|
|[<span data-ttu-id="171d3-151">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="171d3-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="171d3-152">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="171d3-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="171d3-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="171d3-153">SourceProperty :String</span></span>

<span data-ttu-id="171d3-154">Especifica a origem dos dados retornados pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="171d3-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="171d3-155">Tipo:</span><span class="sxs-lookup"><span data-stu-id="171d3-155">Type:</span></span>

*   <span data-ttu-id="171d3-156">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="171d3-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="171d3-157">Properties:</span></span>

|<span data-ttu-id="171d3-158">Nome</span><span class="sxs-lookup"><span data-stu-id="171d3-158">Name</span></span>| <span data-ttu-id="171d3-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="171d3-159">Type</span></span>| <span data-ttu-id="171d3-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="171d3-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="171d3-161">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-161">String</span></span>|<span data-ttu-id="171d3-162">A origem dos dados é do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="171d3-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="171d3-163">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="171d3-163">String</span></span>|<span data-ttu-id="171d3-164">A origem dos dados é do assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="171d3-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="171d3-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="171d3-165">Requirements</span></span>

|<span data-ttu-id="171d3-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="171d3-166">Requirement</span></span>| <span data-ttu-id="171d3-167">Valor</span><span class="sxs-lookup"><span data-stu-id="171d3-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="171d3-168">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="171d3-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="171d3-169">1.0</span><span class="sxs-lookup"><span data-stu-id="171d3-169">1.0</span></span>|
|[<span data-ttu-id="171d3-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="171d3-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="171d3-171">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="171d3-171">Compose or read</span></span>|