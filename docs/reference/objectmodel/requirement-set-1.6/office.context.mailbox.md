
# <a name="mailbox"></a><span data-ttu-id="1d4b9-101">caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-101">mailbox</span></span>

### <span data-ttu-id="1d4b9-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="1d4b9-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d4b9-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-105">Requirements</span></span>

|<span data-ttu-id="1d4b9-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-106">Requirement</span></span>| <span data-ttu-id="1d4b9-107">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-109">1.0</span></span>|
|[<span data-ttu-id="1d4b9-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-111">Restricted</span></span>|
|[<span data-ttu-id="1d4b9-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1d4b9-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-114">Members and methods</span></span>

| <span data-ttu-id="1d4b9-115">Membro</span><span class="sxs-lookup"><span data-stu-id="1d4b9-115">Member</span></span> | <span data-ttu-id="1d4b9-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1d4b9-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="1d4b9-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="1d4b9-118">Membro</span><span class="sxs-lookup"><span data-stu-id="1d4b9-118">Member</span></span> |
| [<span data-ttu-id="1d4b9-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="1d4b9-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="1d4b9-120">Membro</span><span class="sxs-lookup"><span data-stu-id="1d4b9-120">Member</span></span> |
| [<span data-ttu-id="1d4b9-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1d4b9-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1d4b9-122">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-122">Method</span></span> |
| [<span data-ttu-id="1d4b9-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="1d4b9-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="1d4b9-124">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-124">Method</span></span> |
| [<span data-ttu-id="1d4b9-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1d4b9-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="1d4b9-126">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-126">Method</span></span> |
| [<span data-ttu-id="1d4b9-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="1d4b9-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="1d4b9-128">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-128">Method</span></span> |
| [<span data-ttu-id="1d4b9-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="1d4b9-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="1d4b9-130">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-130">Method</span></span> |
| [<span data-ttu-id="1d4b9-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1d4b9-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="1d4b9-132">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-132">Method</span></span> |
| [<span data-ttu-id="1d4b9-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="1d4b9-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="1d4b9-134">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-134">Method</span></span> |
| [<span data-ttu-id="1d4b9-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1d4b9-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="1d4b9-136">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-136">Method</span></span> |
| [<span data-ttu-id="1d4b9-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="1d4b9-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="1d4b9-138">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-138">Method</span></span> |
| [<span data-ttu-id="1d4b9-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d4b9-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="1d4b9-140">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-140">Method</span></span> |
| [<span data-ttu-id="1d4b9-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d4b9-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="1d4b9-142">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-142">Method</span></span> |
| [<span data-ttu-id="1d4b9-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d4b9-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="1d4b9-144">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-144">Method</span></span> |
| [<span data-ttu-id="1d4b9-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="1d4b9-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="1d4b9-146">Método</span><span class="sxs-lookup"><span data-stu-id="1d4b9-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1d4b9-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="1d4b9-147">Namespaces</span></span>

<span data-ttu-id="1d4b9-148">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1d4b9-149">[item](Office.context.mailbox.item.md): Fornece métodos e propriedades para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1d4b9-150">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1d4b9-151">Membros</span><span class="sxs-lookup"><span data-stu-id="1d4b9-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1d4b9-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="1d4b9-152">ewsUrl :String</span></span>

<span data-ttu-id="1d4b9-p102">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-155">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-p103">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1d4b9-158">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="1d4b9-p104">No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1d4b9-161">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-161">Type:</span></span>

*   <span data-ttu-id="1d4b9-162">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d4b9-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-163">Requirements</span></span>

|<span data-ttu-id="1d4b9-164">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-164">Requirement</span></span>| <span data-ttu-id="1d4b9-165">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-166">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-167">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-167">1.0</span></span>|
|[<span data-ttu-id="1d4b9-168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-169">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-171">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="1d4b9-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="1d4b9-172">restUrl :String</span></span>

<span data-ttu-id="1d4b9-173">Obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="1d4b9-174">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](https://docs.microsoft.com/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="1d4b9-175">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="1d4b9-p105">No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="1d4b9-178">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-178">Type:</span></span>

*   <span data-ttu-id="1d4b9-179">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d4b9-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-180">Requirements</span></span>

|<span data-ttu-id="1d4b9-181">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-181">Requirement</span></span>| <span data-ttu-id="1d4b9-182">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-183">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-184">1.5</span><span class="sxs-lookup"><span data-stu-id="1d4b9-184">1.5</span></span> |
|[<span data-ttu-id="1d4b9-185">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-186">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-188">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="1d4b9-189">Métodos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1d4b9-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d4b9-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1d4b9-191">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1d4b9-p106">No momento, o único tipo de evento com suporte é `Office.EventType.ItemChanged`, que é chamado quando o usuário seleciona um novo item. Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface de usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-194">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-194">Parameters:</span></span>

| <span data-ttu-id="1d4b9-195">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-195">Name</span></span> | <span data-ttu-id="1d4b9-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-196">Type</span></span> | <span data-ttu-id="1d4b9-197">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-197">Attributes</span></span> | <span data-ttu-id="1d4b9-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1d4b9-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1d4b9-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1d4b9-200">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1d4b9-201">Função</span><span class="sxs-lookup"><span data-stu-id="1d4b9-201">Function</span></span> || <span data-ttu-id="1d4b9-p107">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1d4b9-205">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-205">Object</span></span> | <span data-ttu-id="1d4b9-206">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-206">&lt;optional&gt;</span></span> | <span data-ttu-id="1d4b9-207">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d4b9-208">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-208">Object</span></span> | <span data-ttu-id="1d4b9-209">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-209">&lt;optional&gt;</span></span> | <span data-ttu-id="1d4b9-210">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1d4b9-211">function</span><span class="sxs-lookup"><span data-stu-id="1d4b9-211">function</span></span>| <span data-ttu-id="1d4b9-212">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-212">&lt;optional&gt;</span></span>|<span data-ttu-id="1d4b9-213">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-214">Requirements</span></span>

|<span data-ttu-id="1d4b9-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-215">Requirement</span></span>| <span data-ttu-id="1d4b9-216">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-217">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-218">1.5</span><span class="sxs-lookup"><span data-stu-id="1d4b9-218">1.5</span></span> |
|[<span data-ttu-id="1d4b9-219">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-220">ReadItem</span></span> |
|[<span data-ttu-id="1d4b9-221">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-222">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-223">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-223">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="1d4b9-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d4b9-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1d4b9-225">Converte uma ID de item formatada para REST em formato EWS.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-226">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-226">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-p108">As IDs de itens recuperadas por meio de uma API REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada em REST no formato adequado para o EWS.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-229">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-229">Parameters:</span></span>

|<span data-ttu-id="1d4b9-230">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-230">Name</span></span>| <span data-ttu-id="1d4b9-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-231">Type</span></span>| <span data-ttu-id="1d4b9-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d4b9-233">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-233">String</span></span>|<span data-ttu-id="1d4b9-234">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="1d4b9-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="1d4b9-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1d4b9-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="1d4b9-236">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-237">Requirements</span></span>

|<span data-ttu-id="1d4b9-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-238">Requirement</span></span>| <span data-ttu-id="1d4b9-239">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-240">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-241">1.3</span><span class="sxs-lookup"><span data-stu-id="1d4b9-241">1.3</span></span>|
|[<span data-ttu-id="1d4b9-242">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-243">Restrito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-243">Restricted</span></span>|
|[<span data-ttu-id="1d4b9-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-245">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d4b9-246">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-246">Returns:</span></span>

<span data-ttu-id="1d4b9-247">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="1d4b9-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1d4b9-248">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="1d4b9-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="1d4b9-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="1d4b9-250">Obtém um dicionário contendo informações da hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1d4b9-p109">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; O Outlook Web App usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve manipular valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário esperado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1d4b9-p110">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-256">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-256">Parameters:</span></span>

|<span data-ttu-id="1d4b9-257">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-257">Name</span></span>| <span data-ttu-id="1d4b9-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-258">Type</span></span>| <span data-ttu-id="1d4b9-259">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1d4b9-260">Date</span><span class="sxs-lookup"><span data-stu-id="1d4b9-260">Date</span></span>|<span data-ttu-id="1d4b9-261">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="1d4b9-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-262">Requirements</span></span>

|<span data-ttu-id="1d4b9-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-263">Requirement</span></span>| <span data-ttu-id="1d4b9-264">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-265">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-266">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-266">1.0</span></span>|
|[<span data-ttu-id="1d4b9-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-268">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-269">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-270">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d4b9-271">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-271">Returns:</span></span>

<span data-ttu-id="1d4b9-272">Tipo: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="1d4b9-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d4b9-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="1d4b9-274">Converte uma ID de item formatada para EWS em formato REST.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-275">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-275">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-p111">As IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-278">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-278">Parameters:</span></span>

|<span data-ttu-id="1d4b9-279">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-279">Name</span></span>| <span data-ttu-id="1d4b9-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-280">Type</span></span>| <span data-ttu-id="1d4b9-281">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d4b9-282">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-282">String</span></span>|<span data-ttu-id="1d4b9-283">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="1d4b9-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="1d4b9-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="1d4b9-285">Um valor indicando a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-286">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-286">Requirements</span></span>

|<span data-ttu-id="1d4b9-287">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-287">Requirement</span></span>| <span data-ttu-id="1d4b9-288">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-289">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-290">1.3</span><span class="sxs-lookup"><span data-stu-id="1d4b9-290">1.3</span></span>|
|[<span data-ttu-id="1d4b9-291">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-292">Restrito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-292">Restricted</span></span>|
|[<span data-ttu-id="1d4b9-293">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-294">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d4b9-295">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-295">Returns:</span></span>

<span data-ttu-id="1d4b9-296">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="1d4b9-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1d4b9-297">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1d4b9-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1d4b9-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1d4b9-299">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1d4b9-300">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-301">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-301">Parameters:</span></span>

|<span data-ttu-id="1d4b9-302">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-302">Name</span></span>| <span data-ttu-id="1d4b9-303">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-303">Type</span></span>| <span data-ttu-id="1d4b9-304">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1d4b9-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1d4b9-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="1d4b9-306">O valor de hora local a ser convertido.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-307">Requirements</span></span>

|<span data-ttu-id="1d4b9-308">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-308">Requirement</span></span>| <span data-ttu-id="1d4b9-309">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-310">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-311">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-311">1.0</span></span>|
|[<span data-ttu-id="1d4b9-312">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-313">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-314">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-315">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d4b9-316">Retorna:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-316">Returns:</span></span>

<span data-ttu-id="1d4b9-317">Um objeto Date com o horário expresso em UTC.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="1d4b9-318">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d4b9-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d4b9-319">Date</span><span class="sxs-lookup"><span data-stu-id="1d4b9-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="1d4b9-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1d4b9-321">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-322">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-322">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-323">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1d4b9-p112">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso principal de uma série recorrente, mas você não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1d4b9-326">No aplicativo Web do Outlook, esse método abre o formulário especificado somente se o corpo do formulário for menor ou igual a um número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1d4b9-327">Se o identificador de item especificado não identificar um compromisso existente, um painel em branco será aberto no computador ou no dispositivo cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-328">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-328">Parameters:</span></span>

|<span data-ttu-id="1d4b9-329">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-329">Name</span></span>| <span data-ttu-id="1d4b9-330">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-330">Type</span></span>| <span data-ttu-id="1d4b9-331">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d4b9-332">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-332">String</span></span>|<span data-ttu-id="1d4b9-333">O identificador de serviços da Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-334">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-334">Requirements</span></span>

|<span data-ttu-id="1d4b9-335">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-335">Requirement</span></span>| <span data-ttu-id="1d4b9-336">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-337">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-338">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-338">1.0</span></span>|
|[<span data-ttu-id="1d4b9-339">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-340">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-341">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-342">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-343">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="1d4b9-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1d4b9-345">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-346">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-346">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-347">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1d4b9-348">No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual a um número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1d4b9-349">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida uma mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1d4b9-p113">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-352">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-352">Parameters:</span></span>

|<span data-ttu-id="1d4b9-353">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-353">Name</span></span>| <span data-ttu-id="1d4b9-354">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-354">Type</span></span>| <span data-ttu-id="1d4b9-355">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1d4b9-356">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-356">String</span></span>|<span data-ttu-id="1d4b9-357">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-358">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-358">Requirements</span></span>

|<span data-ttu-id="1d4b9-359">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-359">Requirement</span></span>| <span data-ttu-id="1d4b9-360">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-361">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-362">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-362">1.0</span></span>|
|[<span data-ttu-id="1d4b9-363">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-364">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-365">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-366">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-367">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1d4b9-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1d4b9-369">Exibe um formulário para criar um novo compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-370">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-370">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d4b9-p114">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1d4b9-p115">No aplicativo Web do Outlook e no OWA para Dispositivos, esse método sempre exibe um formulário com um campo de participantes. Se você não especificar nenhum participante como argumento de entrada, o método exibirá um formulário com um botão **Salvar** . Se você especificar participantes, o formulário incluirá os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1d4b9-p116">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees` ou `resources`, o método exibirá um formulário de reunião com um botão **Enviar** . Se você não especificar destinatários, o método exibirá um formulário de compromisso com um botão **Salvar & Fechar**.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1d4b9-378">Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, será gerada uma exceção.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-379">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-380">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-380">Note: All parameters are optional.</span></span>

|<span data-ttu-id="1d4b9-381">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-381">Name</span></span>| <span data-ttu-id="1d4b9-382">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-382">Type</span></span>| <span data-ttu-id="1d4b9-383">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1d4b9-384">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-384">Object</span></span> | <span data-ttu-id="1d4b9-385">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1d4b9-386">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d4b9-p117">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1d4b9-389">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d4b9-p118">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1d4b9-392">Date</span><span class="sxs-lookup"><span data-stu-id="1d4b9-392">Date</span></span> | <span data-ttu-id="1d4b9-393">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1d4b9-394">Date</span><span class="sxs-lookup"><span data-stu-id="1d4b9-394">Date</span></span> | <span data-ttu-id="1d4b9-395">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1d4b9-396">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-396">String</span></span> | <span data-ttu-id="1d4b9-p119">Uma sequência de caracteres que contém o local do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1d4b9-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1d4b9-p120">Uma matriz de sequências de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1d4b9-402">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-402">String</span></span> | <span data-ttu-id="1d4b9-p121">Uma sequência de caracteres que contém o assunto do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1d4b9-405">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-405">String</span></span> | <span data-ttu-id="1d4b9-p122">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1d4b9-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-408">Requirements</span></span>

|<span data-ttu-id="1d4b9-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-409">Requirement</span></span>| <span data-ttu-id="1d4b9-410">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-411">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-412">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-412">1.0</span></span>|
|[<span data-ttu-id="1d4b9-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-414">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-416">Leitura</span><span class="sxs-lookup"><span data-stu-id="1d4b9-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-417">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="1d4b9-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="1d4b9-419">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="1d4b9-420">O método `displayNewMessageForm` abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="1d4b9-421">Quando os parâmetros são especificados, os campos do formulário de mensagem são preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1d4b9-422">Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, uma exceção será gerada.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-423">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-424">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-424">Note: All parameters are optional.</span></span>

|<span data-ttu-id="1d4b9-425">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-425">Name</span></span>| <span data-ttu-id="1d4b9-426">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-426">Type</span></span>| <span data-ttu-id="1d4b9-427">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1d4b9-428">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-428">Object</span></span> | <span data-ttu-id="1d4b9-429">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="1d4b9-430">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d4b9-431">Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Para.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="1d4b9-432">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="1d4b9-433">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d4b9-434">Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cc.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="1d4b9-435">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="1d4b9-436">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="1d4b9-437">Uma matriz de sequência de caracteres que contém os endereços de e-mail ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="1d4b9-438">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1d4b9-439">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-439">String</span></span> | <span data-ttu-id="1d4b9-440">Uma sequência de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-440">A string containing the subject of the message.</span></span> <span data-ttu-id="1d4b9-441">A sequência de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="1d4b9-442">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-442">String</span></span> | <span data-ttu-id="1d4b9-443">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-443">The HTML body of the message.</span></span> <span data-ttu-id="1d4b9-444">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="1d4b9-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1d4b9-446">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="1d4b9-447">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-447">String</span></span> | <span data-ttu-id="1d4b9-p129">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="1d4b9-450">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-450">String</span></span> | <span data-ttu-id="1d4b9-451">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="1d4b9-452">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-452">String</span></span> | <span data-ttu-id="1d4b9-p130">Usado somente se `type` estiver definido como `file`. A URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="1d4b9-455">Booleano</span><span class="sxs-lookup"><span data-stu-id="1d4b9-455">Boolean</span></span> | <span data-ttu-id="1d4b9-p131">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="1d4b9-458">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-458">String</span></span> | <span data-ttu-id="1d4b9-459">Usado somente se `type` estiver definido para `item`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="1d4b9-460">A ID do item do EWS do e-mail existente que deseja anexar na nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="1d4b9-461">É uma sequência de caracteres com até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="1d4b9-462">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-462">Requirements</span></span>

|<span data-ttu-id="1d4b9-463">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-463">Requirement</span></span>| <span data-ttu-id="1d4b9-464">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-465">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-465">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-466">1.6</span><span class="sxs-lookup"><span data-stu-id="1d4b9-466">-16</span></span> |
|[<span data-ttu-id="1d4b9-467">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-468">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-469">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-470">Leitura</span><span class="sxs-lookup"><span data-stu-id="1d4b9-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-471">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-471">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="1d4b9-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1d4b9-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="1d4b9-473">Obtém uma sequência de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="1d4b9-p133">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-476">É recomendável que suplementos usem as APIs REST em vez de Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-476">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="1d4b9-477">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="1d4b9-477">**REST Tokens**</span></span>

<span data-ttu-id="1d4b9-p134">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a e-mail, calendário e contatos, incluindo a capacidade de enviar e-mails.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="1d4b9-481">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="1d4b9-482">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="1d4b9-482">**EWS Tokens**</span></span>

<span data-ttu-id="1d4b9-p135">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="1d4b9-485">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-486">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-486">Parameters:</span></span>

|<span data-ttu-id="1d4b9-487">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-487">Name</span></span>| <span data-ttu-id="1d4b9-488">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-488">Type</span></span>| <span data-ttu-id="1d4b9-489">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-489">Attributes</span></span>| <span data-ttu-id="1d4b9-490">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="1d4b9-491">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-491">Object</span></span> | <span data-ttu-id="1d4b9-492">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-492">&lt;optional&gt;</span></span> | <span data-ttu-id="1d4b9-493">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="1d4b9-494">Booleano</span><span class="sxs-lookup"><span data-stu-id="1d4b9-494">Boolean</span></span> |  <span data-ttu-id="1d4b9-495">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-495">&lt;optional&gt;</span></span> | <span data-ttu-id="1d4b9-p136">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d4b9-498">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-498">Object</span></span> |  <span data-ttu-id="1d4b9-499">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-499">&lt;optional&gt;</span></span> | <span data-ttu-id="1d4b9-500">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="1d4b9-501">function</span><span class="sxs-lookup"><span data-stu-id="1d4b9-501">function</span></span>||<span data-ttu-id="1d4b9-p137">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-504">Requirements</span></span>

|<span data-ttu-id="1d4b9-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-505">Requirement</span></span>| <span data-ttu-id="1d4b9-506">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-507">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-508">1.5</span><span class="sxs-lookup"><span data-stu-id="1d4b9-508">1.5</span></span> |
|[<span data-ttu-id="1d4b9-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-510">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-512">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-513">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1d4b9-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d4b9-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d4b9-515">Obtém uma sequência de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1d4b9-p138">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1d4b9-p139">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1d4b9-521">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="1d4b9-p140">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-524">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-524">Parameters:</span></span>

|<span data-ttu-id="1d4b9-525">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-525">Name</span></span>| <span data-ttu-id="1d4b9-526">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-526">Type</span></span>| <span data-ttu-id="1d4b9-527">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-527">Attributes</span></span>| <span data-ttu-id="1d4b9-528">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1d4b9-529">function</span><span class="sxs-lookup"><span data-stu-id="1d4b9-529">function</span></span>||<span data-ttu-id="1d4b9-p141">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1d4b9-532">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-532">Object</span></span>| <span data-ttu-id="1d4b9-533">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-533">&lt;optional&gt;</span></span>|<span data-ttu-id="1d4b9-534">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-535">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-535">Requirements</span></span>

|<span data-ttu-id="1d4b9-536">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-536">Requirement</span></span>| <span data-ttu-id="1d4b9-537">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-538">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-539">1.3</span><span class="sxs-lookup"><span data-stu-id="1d4b9-539">1.3</span></span>|
|[<span data-ttu-id="1d4b9-540">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-541">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-542">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-543">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-544">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1d4b9-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d4b9-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d4b9-546">Obtém um token que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1d4b9-547">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-548">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-548">Parameters:</span></span>

|<span data-ttu-id="1d4b9-549">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-549">Name</span></span>| <span data-ttu-id="1d4b9-550">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-550">Type</span></span>| <span data-ttu-id="1d4b9-551">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-551">Attributes</span></span>| <span data-ttu-id="1d4b9-552">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1d4b9-553">function</span><span class="sxs-lookup"><span data-stu-id="1d4b9-553">function</span></span>||<span data-ttu-id="1d4b9-554">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d4b9-555">O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="1d4b9-556">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-556">Object</span></span>| <span data-ttu-id="1d4b9-557">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-557">&lt;optional&gt;</span></span>|<span data-ttu-id="1d4b9-558">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-559">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-559">Requirements</span></span>

|<span data-ttu-id="1d4b9-560">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-560">Requirement</span></span>| <span data-ttu-id="1d4b9-561">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-562">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-563">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-563">1.0</span></span>|
|[<span data-ttu-id="1d4b9-564">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d4b9-565">ReadItem</span></span>|
|[<span data-ttu-id="1d4b9-566">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-567">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-568">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1d4b9-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d4b9-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1d4b9-570">Faz uma solicitação assíncrona em um dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-571">Esse método não é suportado nos seguintes cenários.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1d4b9-572">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="1d4b9-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="1d4b9-573">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="1d4b9-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1d4b9-574">Nesses casos, os suplementos devem [usar APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1d4b9-575">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="1d4b9-576">Para obter uma lista de operações EWS compatíveis, confira [Chamar serviços Web de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="1d4b9-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1d4b9-577">Não é possível solicitar os itens associados à pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1d4b9-578">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1d4b9-p143">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para obter mais informações sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso do suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1d4b9-581">O administrador do servidor deve definir `OAuthAuthentication` como verdadeiro no diretório EWS do Servidor de Acesso para Cliente para ativar o método `makeEwsRequestAsync` para fazer solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-581">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1d4b9-582">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="1d4b9-582">Version differences</span></span>

<span data-ttu-id="1d4b9-583">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1d4b9-p144">Você não precisa definir o valor de codificação quando seu aplicativo de email estiver sendo executado no Outlook na Web. Você pode determinar se o seu aplicativo de email está sendo executado no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar qual versão do Outlook está sendo executada usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d4b9-587">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="1d4b9-587">Parameters:</span></span>

|<span data-ttu-id="1d4b9-588">Nome</span><span class="sxs-lookup"><span data-stu-id="1d4b9-588">Name</span></span>| <span data-ttu-id="1d4b9-589">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-589">Type</span></span>| <span data-ttu-id="1d4b9-590">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-590">Attributes</span></span>| <span data-ttu-id="1d4b9-591">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d4b9-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1d4b9-592">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d4b9-592">String</span></span>||<span data-ttu-id="1d4b9-593">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1d4b9-594">function</span><span class="sxs-lookup"><span data-stu-id="1d4b9-594">function</span></span>||<span data-ttu-id="1d4b9-595">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d4b9-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d4b9-596">O resultado XML da chamada do EWS é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="1d4b9-597">Se o resultado exceder 1 MB, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-597">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="1d4b9-598">Objeto</span><span class="sxs-lookup"><span data-stu-id="1d4b9-598">Object</span></span>| <span data-ttu-id="1d4b9-599">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="1d4b9-599">&lt;optional&gt;</span></span>|<span data-ttu-id="1d4b9-600">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d4b9-601">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d4b9-601">Requirements</span></span>

|<span data-ttu-id="1d4b9-602">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d4b9-602">Requirement</span></span>| <span data-ttu-id="1d4b9-603">Valor</span><span class="sxs-lookup"><span data-stu-id="1d4b9-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d4b9-604">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d4b9-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d4b9-605">1.0</span><span class="sxs-lookup"><span data-stu-id="1d4b9-605">1.0</span></span>|
|[<span data-ttu-id="1d4b9-606">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d4b9-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1d4b9-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1d4b9-608">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d4b9-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d4b9-609">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1d4b9-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d4b9-610">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d4b9-610">Example</span></span>

<span data-ttu-id="1d4b9-611">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="1d4b9-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```