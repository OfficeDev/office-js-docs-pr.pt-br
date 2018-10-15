
# <a name="mailbox"></a><span data-ttu-id="0d8d4-101">caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-101">mailbox</span></span>

### <span data-ttu-id="0d8d4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="0d8d4-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d8d4-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-105">Requirements</span></span>

|<span data-ttu-id="0d8d4-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-106">Requirement</span></span>| <span data-ttu-id="0d8d4-107">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-109">1.0</span></span>|
|[<span data-ttu-id="0d8d4-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-111">Restricted</span></span>|
|[<span data-ttu-id="0d8d4-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0d8d4-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-114">Members and methods</span></span>

| <span data-ttu-id="0d8d4-115">Membro</span><span class="sxs-lookup"><span data-stu-id="0d8d4-115">Member</span></span> | <span data-ttu-id="0d8d4-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0d8d4-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="0d8d4-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="0d8d4-118">Membro</span><span class="sxs-lookup"><span data-stu-id="0d8d4-118">Member</span></span> |
| [<span data-ttu-id="0d8d4-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="0d8d4-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="0d8d4-120">Membro</span><span class="sxs-lookup"><span data-stu-id="0d8d4-120">Member</span></span> |
| [<span data-ttu-id="0d8d4-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0d8d4-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0d8d4-122">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-122">Method</span></span> |
| [<span data-ttu-id="0d8d4-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="0d8d4-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="0d8d4-124">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-124">Method</span></span> |
| [<span data-ttu-id="0d8d4-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0d8d4-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="0d8d4-126">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-126">Method</span></span> |
| [<span data-ttu-id="0d8d4-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="0d8d4-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="0d8d4-128">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-128">Method</span></span> |
| [<span data-ttu-id="0d8d4-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="0d8d4-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="0d8d4-130">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-130">Method</span></span> |
| [<span data-ttu-id="0d8d4-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0d8d4-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="0d8d4-132">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-132">Method</span></span> |
| [<span data-ttu-id="0d8d4-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="0d8d4-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="0d8d4-134">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-134">Method</span></span> |
| [<span data-ttu-id="0d8d4-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0d8d4-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="0d8d4-136">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-136">Method</span></span> |
| [<span data-ttu-id="0d8d4-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0d8d4-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="0d8d4-138">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-138">Method</span></span> |
| [<span data-ttu-id="0d8d4-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0d8d4-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="0d8d4-140">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-140">Method</span></span> |
| [<span data-ttu-id="0d8d4-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0d8d4-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="0d8d4-142">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-142">Method</span></span> |
| [<span data-ttu-id="0d8d4-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="0d8d4-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="0d8d4-144">Método</span><span class="sxs-lookup"><span data-stu-id="0d8d4-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0d8d4-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="0d8d4-145">Namespaces</span></span>

<span data-ttu-id="0d8d4-146">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="0d8d4-147">[item](Office.context.mailbox.item.md): Fornece métodos e propriedades para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="0d8d4-148">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="0d8d4-149">Membros</span><span class="sxs-lookup"><span data-stu-id="0d8d4-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="0d8d4-150">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="0d8d4-150">ewsUrl :String</span></span>

<span data-ttu-id="0d8d4-p102">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-153">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-153">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-p103">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0d8d4-156">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="0d8d4-p104">No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0d8d4-159">Tipo:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-159">Type:</span></span>

*   <span data-ttu-id="0d8d4-160">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d8d4-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-161">Requirements</span></span>

|<span data-ttu-id="0d8d4-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-162">Requirement</span></span>| <span data-ttu-id="0d8d4-163">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-164">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-165">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-165">1.0</span></span>|
|[<span data-ttu-id="0d8d4-166">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-167">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-168">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-169">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="0d8d4-170">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="0d8d4-170">restUrl :String</span></span>

<span data-ttu-id="0d8d4-171">Obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="0d8d4-172">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](https://docs.microsoft.com/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="0d8d4-173">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="0d8d4-p105">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-176">Clientes Outlook conectados a instalações locais do Exchange 2016 ou posterior com uma URL REST personalizada configurada retornarão um valor inválido para `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-176">Note: Outlook clients connected to on-premises installations of Exchange 2016 with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d8d4-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-177">Type:</span></span>

*   <span data-ttu-id="0d8d4-178">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d8d4-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-179">Requirements</span></span>

|<span data-ttu-id="0d8d4-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-180">Requirement</span></span>| <span data-ttu-id="0d8d4-181">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-182">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-183">1.5</span><span class="sxs-lookup"><span data-stu-id="0d8d4-183">1.5</span></span> |
|[<span data-ttu-id="0d8d4-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-185">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-187">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="0d8d4-188">Métodos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0d8d4-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d8d4-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0d8d4-190">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0d8d4-p106">No momento, o único tipo de evento com suporte é `Office.EventType.ItemChanged`, que é chamado quando o usuário seleciona um novo item. Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface de usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-193">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-193">Parameters:</span></span>

| <span data-ttu-id="0d8d4-194">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-194">Name</span></span> | <span data-ttu-id="0d8d4-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-195">Type</span></span> | <span data-ttu-id="0d8d4-196">Atributos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-196">Attributes</span></span> | <span data-ttu-id="0d8d4-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0d8d4-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0d8d4-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0d8d4-199">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0d8d4-200">Função</span><span class="sxs-lookup"><span data-stu-id="0d8d4-200">Function</span></span> || <span data-ttu-id="0d8d4-p107">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0d8d4-204">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-204">Object</span></span> | <span data-ttu-id="0d8d4-205">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-205">&lt;optional&gt;</span></span> | <span data-ttu-id="0d8d4-206">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0d8d4-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-207">Object</span></span> | <span data-ttu-id="0d8d4-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-208">&lt;optional&gt;</span></span> | <span data-ttu-id="0d8d4-209">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0d8d4-210">function</span><span class="sxs-lookup"><span data-stu-id="0d8d4-210">function</span></span>| <span data-ttu-id="0d8d4-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-211">&lt;optional&gt;</span></span>|<span data-ttu-id="0d8d4-212">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-213">Requirements</span></span>

|<span data-ttu-id="0d8d4-214">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-214">Requirement</span></span>| <span data-ttu-id="0d8d4-215">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-216">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-217">1.5</span><span class="sxs-lookup"><span data-stu-id="0d8d4-217">1.5</span></span> |
|[<span data-ttu-id="0d8d4-218">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-219">ReadItem</span></span> |
|[<span data-ttu-id="0d8d4-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-221">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="0d8d4-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0d8d4-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0d8d4-224">Converte uma ID de item formatada para REST em formato EWS.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-225">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-p108">As IDs de itens recuperadas por meio de uma API REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada em REST no formato adequado para o EWS.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-228">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-228">Parameters:</span></span>

|<span data-ttu-id="0d8d4-229">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-229">Name</span></span>| <span data-ttu-id="0d8d4-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-230">Type</span></span>| <span data-ttu-id="0d8d4-231">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0d8d4-232">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-232">String</span></span>|<span data-ttu-id="0d8d4-233">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="0d8d4-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="0d8d4-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0d8d4-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="0d8d4-235">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-236">Requirements</span></span>

|<span data-ttu-id="0d8d4-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-237">Requirement</span></span>| <span data-ttu-id="0d8d4-238">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-239">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-240">1.3</span><span class="sxs-lookup"><span data-stu-id="0d8d4-240">1.3</span></span>|
|[<span data-ttu-id="0d8d4-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-242">Restrito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-242">Restricted</span></span>|
|[<span data-ttu-id="0d8d4-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-244">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d8d4-245">Retorna:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-245">Returns:</span></span>

<span data-ttu-id="0d8d4-246">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="0d8d4-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0d8d4-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="0d8d4-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="0d8d4-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="0d8d4-249">Obtém um dicionário contendo informações da hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="0d8d4-p109">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; O Outlook Web App usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve manipular valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário esperado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="0d8d4-p110">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-255">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-255">Parameters:</span></span>

|<span data-ttu-id="0d8d4-256">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-256">Name</span></span>| <span data-ttu-id="0d8d4-257">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-257">Type</span></span>| <span data-ttu-id="0d8d4-258">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="0d8d4-259">Date</span><span class="sxs-lookup"><span data-stu-id="0d8d4-259">Date</span></span>|<span data-ttu-id="0d8d4-260">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="0d8d4-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-261">Requirements</span></span>

|<span data-ttu-id="0d8d4-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-262">Requirement</span></span>| <span data-ttu-id="0d8d4-263">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-264">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-265">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-265">1.0</span></span>|
|[<span data-ttu-id="0d8d4-266">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-267">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-268">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-269">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d8d4-270">Retorna:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-270">Returns:</span></span>

<span data-ttu-id="0d8d4-271">Tipo: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="0d8d4-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0d8d4-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0d8d4-273">Converte uma ID de item formatada para EWS em formato REST.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-274">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-p111">As IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-277">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-277">Parameters:</span></span>

|<span data-ttu-id="0d8d4-278">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-278">Name</span></span>| <span data-ttu-id="0d8d4-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-279">Type</span></span>| <span data-ttu-id="0d8d4-280">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0d8d4-281">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-281">String</span></span>|<span data-ttu-id="0d8d4-282">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="0d8d4-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0d8d4-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="0d8d4-284">Um valor indicando a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-285">Requirements</span></span>

|<span data-ttu-id="0d8d4-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-286">Requirement</span></span>| <span data-ttu-id="0d8d4-287">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-288">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-289">1.3</span><span class="sxs-lookup"><span data-stu-id="0d8d4-289">1.3</span></span>|
|[<span data-ttu-id="0d8d4-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-291">Restrito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-291">Restricted</span></span>|
|[<span data-ttu-id="0d8d4-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-293">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d8d4-294">Retorna:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-294">Returns:</span></span>

<span data-ttu-id="0d8d4-295">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="0d8d4-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0d8d4-296">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="0d8d4-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="0d8d4-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="0d8d4-298">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="0d8d4-299">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-300">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-300">Parameters:</span></span>

|<span data-ttu-id="0d8d4-301">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-301">Name</span></span>| <span data-ttu-id="0d8d4-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-302">Type</span></span>| <span data-ttu-id="0d8d4-303">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="0d8d4-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0d8d4-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="0d8d4-305">O valor de hora local a ser convertido.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-306">Requirements</span></span>

|<span data-ttu-id="0d8d4-307">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-307">Requirement</span></span>| <span data-ttu-id="0d8d4-308">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-309">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-310">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-310">1.0</span></span>|
|[<span data-ttu-id="0d8d4-311">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-312">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-314">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d8d4-315">Retorna:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-315">Returns:</span></span>

<span data-ttu-id="0d8d4-316">Um objeto Date com o horário expresso em UTC.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="0d8d4-317">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="0d8d4-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0d8d4-318">Date</span><span class="sxs-lookup"><span data-stu-id="0d8d4-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="0d8d4-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="0d8d4-320">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-321">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-322">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0d8d4-p112">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso principal de uma série recorrente, mas você não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="0d8d4-325">No aplicativo Web do Outlook, esse método abre o formulário especificado somente se o corpo do formulário for menor ou igual a um número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="0d8d4-326">Se o identificador de item especificado não identificar um compromisso existente, um painel em branco será aberto no computador ou no dispositivo cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-327">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-327">Parameters:</span></span>

|<span data-ttu-id="0d8d4-328">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-328">Name</span></span>| <span data-ttu-id="0d8d4-329">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-329">Type</span></span>| <span data-ttu-id="0d8d4-330">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0d8d4-331">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-331">String</span></span>|<span data-ttu-id="0d8d4-332">O identificador de serviços da Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-333">Requirements</span></span>

|<span data-ttu-id="0d8d4-334">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-334">Requirement</span></span>| <span data-ttu-id="0d8d4-335">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-336">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-337">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-337">1.0</span></span>|
|[<span data-ttu-id="0d8d4-338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-339">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-341">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-342">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="0d8d4-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="0d8d4-344">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-345">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-346">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0d8d4-347">No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual a um número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="0d8d4-348">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida uma mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="0d8d4-p113">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-351">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-351">Parameters:</span></span>

|<span data-ttu-id="0d8d4-352">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-352">Name</span></span>| <span data-ttu-id="0d8d4-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-353">Type</span></span>| <span data-ttu-id="0d8d4-354">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0d8d4-355">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-355">String</span></span>|<span data-ttu-id="0d8d4-356">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-357">Requirements</span></span>

|<span data-ttu-id="0d8d4-358">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-358">Requirement</span></span>| <span data-ttu-id="0d8d4-359">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-360">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-361">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-361">1.0</span></span>|
|[<span data-ttu-id="0d8d4-362">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-363">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-364">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-365">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-366">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="0d8d4-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="0d8d4-368">Exibe um formulário para criar um novo compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-369">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0d8d4-p114">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0d8d4-p115">No aplicativo Web do Outlook e no OWA para Dispositivos, esse método sempre exibe um formulário com um campo de participantes. Se você não especificar nenhum participante como argumento de entrada, o método exibirá um formulário com um botão **Salvar** . Se você especificar participantes, o formulário incluirá os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="0d8d4-p116">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees` ou `resources`, o método exibirá um formulário de reunião com um botão **Enviar** . Se você não especificar destinatários, o método exibirá um formulário de compromisso com um botão **Salvar & Fechar**.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="0d8d4-377">Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, será gerada uma exceção.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-378">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-378">Parameters:</span></span>

|<span data-ttu-id="0d8d4-379">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-379">Name</span></span>| <span data-ttu-id="0d8d4-380">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-380">Type</span></span>| <span data-ttu-id="0d8d4-381">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0d8d4-382">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-382">Object</span></span> | <span data-ttu-id="0d8d4-383">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="0d8d4-384">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0d8d4-p117">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="0d8d4-387">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0d8d4-p118">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="0d8d4-390">Date</span><span class="sxs-lookup"><span data-stu-id="0d8d4-390">Date</span></span> | <span data-ttu-id="0d8d4-391">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="0d8d4-392">Date</span><span class="sxs-lookup"><span data-stu-id="0d8d4-392">Date</span></span> | <span data-ttu-id="0d8d4-393">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="0d8d4-394">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-394">String</span></span> | <span data-ttu-id="0d8d4-p119">Uma sequência de caracteres que contém o local do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="0d8d4-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="0d8d4-p120">Uma matriz de sequências de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0d8d4-400">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-400">String</span></span> | <span data-ttu-id="0d8d4-p121">Uma sequência de caracteres que contém o assunto do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="0d8d4-403">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-403">String</span></span> | <span data-ttu-id="0d8d4-p122">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d8d4-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-406">Requirements</span></span>

|<span data-ttu-id="0d8d4-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-407">Requirement</span></span>| <span data-ttu-id="0d8d4-408">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-409">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-410">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-410">1.0</span></span>|
|[<span data-ttu-id="0d8d4-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-412">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-414">Leitura</span><span class="sxs-lookup"><span data-stu-id="0d8d4-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-415">Example</span></span>

```js
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="0d8d4-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0d8d4-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="0d8d4-417">Obtém uma sequência de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="0d8d4-p123">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-420">É recomendável que suplementos usem as APIs REST em vez de Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-420">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="0d8d4-421">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="0d8d4-421">**REST Tokens**</span></span>

<span data-ttu-id="0d8d4-p124">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a e-mail, calendário e contatos, incluindo a capacidade de enviar e-mails.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="0d8d4-425">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="0d8d4-426">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="0d8d4-426">**EWS Tokens**</span></span>

<span data-ttu-id="0d8d4-p125">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="0d8d4-429">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-430">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-430">Parameters:</span></span>

|<span data-ttu-id="0d8d4-431">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-431">Name</span></span>| <span data-ttu-id="0d8d4-432">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-432">Type</span></span>| <span data-ttu-id="0d8d4-433">Atributos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-433">Attributes</span></span>| <span data-ttu-id="0d8d4-434">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="0d8d4-435">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-435">Object</span></span> | <span data-ttu-id="0d8d4-436">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-436">&lt;optional&gt;</span></span> | <span data-ttu-id="0d8d4-437">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="0d8d4-438">Booleano</span><span class="sxs-lookup"><span data-stu-id="0d8d4-438">Boolean</span></span> |  <span data-ttu-id="0d8d4-439">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-439">&lt;optional&gt;</span></span> | <span data-ttu-id="0d8d4-p126">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0d8d4-442">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-442">Object</span></span> |  <span data-ttu-id="0d8d4-443">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-443">&lt;optional&gt;</span></span> | <span data-ttu-id="0d8d4-444">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="0d8d4-445">function</span><span class="sxs-lookup"><span data-stu-id="0d8d4-445">function</span></span>||<span data-ttu-id="0d8d4-p127">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-448">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-448">Requirements</span></span>

|<span data-ttu-id="0d8d4-449">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-449">Requirement</span></span>| <span data-ttu-id="0d8d4-450">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-451">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-452">1.5</span><span class="sxs-lookup"><span data-stu-id="0d8d4-452">1.5</span></span> |
|[<span data-ttu-id="0d8d4-453">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-454">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-455">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-456">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-457">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="0d8d4-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0d8d4-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0d8d4-459">Obtém uma sequência de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="0d8d4-p128">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="0d8d4-p129">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0d8d4-465">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="0d8d4-p130">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-468">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-468">Parameters:</span></span>

|<span data-ttu-id="0d8d4-469">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-469">Name</span></span>| <span data-ttu-id="0d8d4-470">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-470">Type</span></span>| <span data-ttu-id="0d8d4-471">Atributos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-471">Attributes</span></span>| <span data-ttu-id="0d8d4-472">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0d8d4-473">function</span><span class="sxs-lookup"><span data-stu-id="0d8d4-473">function</span></span>||<span data-ttu-id="0d8d4-p131">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0d8d4-476">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-476">Object</span></span>| <span data-ttu-id="0d8d4-477">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-477">&lt;optional&gt;</span></span>|<span data-ttu-id="0d8d4-478">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-479">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-479">Requirements</span></span>

|<span data-ttu-id="0d8d4-480">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-480">Requirement</span></span>| <span data-ttu-id="0d8d4-481">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-482">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-483">1.3</span><span class="sxs-lookup"><span data-stu-id="0d8d4-483">1.3</span></span>|
|[<span data-ttu-id="0d8d4-484">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-485">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-486">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-487">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-488">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="0d8d4-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0d8d4-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0d8d4-490">Obtém um token que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="0d8d4-491">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-492">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-492">Parameters:</span></span>

|<span data-ttu-id="0d8d4-493">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-493">Name</span></span>| <span data-ttu-id="0d8d4-494">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-494">Type</span></span>| <span data-ttu-id="0d8d4-495">Atributos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-495">Attributes</span></span>| <span data-ttu-id="0d8d4-496">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0d8d4-497">function</span><span class="sxs-lookup"><span data-stu-id="0d8d4-497">function</span></span>||<span data-ttu-id="0d8d4-498">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d8d4-499">O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0d8d4-500">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-500">Object</span></span>| <span data-ttu-id="0d8d4-501">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-501">&lt;optional&gt;</span></span>|<span data-ttu-id="0d8d4-502">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-503">Requirements</span></span>

|<span data-ttu-id="0d8d4-504">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-504">Requirement</span></span>| <span data-ttu-id="0d8d4-505">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-506">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-507">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-507">1.0</span></span>|
|[<span data-ttu-id="0d8d4-508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d8d4-509">ReadItem</span></span>|
|[<span data-ttu-id="0d8d4-510">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-511">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-512">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="0d8d4-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0d8d4-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="0d8d4-514">Faz uma solicitação assíncrona em um dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-515">Esse método não é suportado nos seguintes cenários.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="0d8d4-516">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="0d8d4-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="0d8d4-517">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="0d8d4-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="0d8d4-518">Nesses casos, os suplementos devem [usar APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="0d8d4-p132">O método `makeEwsRequestAsync` envia uma solicitação EWS em nome do suplemento para o Exchange. Para obter uma lista das operações EWS suportadas, consulte [Chamar serviços web a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p132">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="0d8d4-521">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="0d8d4-522">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="0d8d4-p133">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para obter mais informações sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso do suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="0d8d4-525">O administrador do servidor deve definir `OAuthAuthentication` como verdadeiro no diretório EWS do Servidor de Acesso para Cliente para ativar o método `makeEwsRequestAsync` para fazer solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-525">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="0d8d4-526">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="0d8d4-526">Version differences</span></span>

<span data-ttu-id="0d8d4-527">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="0d8d4-p134">Você não precisa definir o valor de codificação quando seu aplicativo de email estiver sendo executado no Outlook na Web. Você pode determinar se o seu aplicativo de email está sendo executado no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar qual versão do Outlook está sendo executada usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d8d4-531">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="0d8d4-531">Parameters:</span></span>

|<span data-ttu-id="0d8d4-532">Nome</span><span class="sxs-lookup"><span data-stu-id="0d8d4-532">Name</span></span>| <span data-ttu-id="0d8d4-533">Tipo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-533">Type</span></span>| <span data-ttu-id="0d8d4-534">Atributos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-534">Attributes</span></span>| <span data-ttu-id="0d8d4-535">Descrição</span><span class="sxs-lookup"><span data-stu-id="0d8d4-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0d8d4-536">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="0d8d4-536">String</span></span>||<span data-ttu-id="0d8d4-537">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="0d8d4-538">function</span><span class="sxs-lookup"><span data-stu-id="0d8d4-538">function</span></span>||<span data-ttu-id="0d8d4-539">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0d8d4-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d8d4-p135">O resultado XML da chamada EWS é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`. Se o resultado exceder 1 MB em tamanho, uma mensagem de erro será retornada em vez disso.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-p135">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="0d8d4-542">Objeto</span><span class="sxs-lookup"><span data-stu-id="0d8d4-542">Object</span></span>| <span data-ttu-id="0d8d4-543">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d8d4-543">&lt;optional&gt;</span></span>|<span data-ttu-id="0d8d4-544">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d8d4-545">Requisitos</span><span class="sxs-lookup"><span data-stu-id="0d8d4-545">Requirements</span></span>

|<span data-ttu-id="0d8d4-546">Requisito</span><span class="sxs-lookup"><span data-stu-id="0d8d4-546">Requirement</span></span>| <span data-ttu-id="0d8d4-547">Valor</span><span class="sxs-lookup"><span data-stu-id="0d8d4-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d8d4-548">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="0d8d4-548">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d8d4-549">1.0</span><span class="sxs-lookup"><span data-stu-id="0d8d4-549">1.0</span></span>|
|[<span data-ttu-id="0d8d4-550">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d8d4-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="0d8d4-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="0d8d4-552">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="0d8d4-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0d8d4-553">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="0d8d4-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d8d4-554">Exemplo</span><span class="sxs-lookup"><span data-stu-id="0d8d4-554">Example</span></span>

<span data-ttu-id="0d8d4-555">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="0d8d4-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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