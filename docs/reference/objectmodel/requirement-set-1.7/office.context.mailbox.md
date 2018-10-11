
# <a name="mailbox"></a><span data-ttu-id="31dd9-101">caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-101">mailbox</span></span>

### <span data-ttu-id="31dd9-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="31dd9-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="31dd9-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="31dd9-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31dd9-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-105">Requirements</span></span>

|<span data-ttu-id="31dd9-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-106">Requirement</span></span>| <span data-ttu-id="31dd9-107">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-109">1.0</span></span>|
|[<span data-ttu-id="31dd9-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="31dd9-111">Restricted</span></span>|
|[<span data-ttu-id="31dd9-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="31dd9-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="31dd9-114">Members and methods</span></span>

| <span data-ttu-id="31dd9-115">Membro</span><span class="sxs-lookup"><span data-stu-id="31dd9-115">Member</span></span> | <span data-ttu-id="31dd9-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="31dd9-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="31dd9-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="31dd9-118">Membro</span><span class="sxs-lookup"><span data-stu-id="31dd9-118">Member</span></span> |
| [<span data-ttu-id="31dd9-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="31dd9-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="31dd9-120">Membro</span><span class="sxs-lookup"><span data-stu-id="31dd9-120">Member</span></span> |
| [<span data-ttu-id="31dd9-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="31dd9-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="31dd9-122">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-122">Method</span></span> |
| [<span data-ttu-id="31dd9-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="31dd9-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="31dd9-124">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-124">Method</span></span> |
| [<span data-ttu-id="31dd9-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="31dd9-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="31dd9-126">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-126">Method</span></span> |
| [<span data-ttu-id="31dd9-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="31dd9-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="31dd9-128">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-128">Method</span></span> |
| [<span data-ttu-id="31dd9-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="31dd9-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="31dd9-130">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-130">Method</span></span> |
| [<span data-ttu-id="31dd9-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="31dd9-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="31dd9-132">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-132">Method</span></span> |
| [<span data-ttu-id="31dd9-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="31dd9-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="31dd9-134">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-134">Method</span></span> |
| [<span data-ttu-id="31dd9-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="31dd9-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="31dd9-136">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-136">Method</span></span> |
| [<span data-ttu-id="31dd9-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="31dd9-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="31dd9-138">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-138">Method</span></span> |
| [<span data-ttu-id="31dd9-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="31dd9-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="31dd9-140">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-140">Method</span></span> |
| [<span data-ttu-id="31dd9-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="31dd9-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="31dd9-142">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-142">Method</span></span> |
| [<span data-ttu-id="31dd9-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="31dd9-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="31dd9-144">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-144">Method</span></span> |
| [<span data-ttu-id="31dd9-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="31dd9-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="31dd9-146">Método</span><span class="sxs-lookup"><span data-stu-id="31dd9-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="31dd9-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="31dd9-147">Namespaces</span></span>

<span data-ttu-id="31dd9-148">[diagnostics](Office.context.mailbox.diagnostics.md): fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="31dd9-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="31dd9-149">[item](Office.context.mailbox.item.md): fornece métodos e propriedades para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="31dd9-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="31dd9-150">[userProfile](Office.context.mailbox.userProfile.md): fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="31dd9-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="31dd9-151">Membros</span><span class="sxs-lookup"><span data-stu-id="31dd9-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="31dd9-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="31dd9-152">ewsUrl :String</span></span>

<span data-ttu-id="31dd9-p102">Obtém o URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-155">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-p103">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="31dd9-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="31dd9-158">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31dd9-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="31dd9-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="31dd9-161">Tipo:</span><span class="sxs-lookup"><span data-stu-id="31dd9-161">Type:</span></span>

*   <span data-ttu-id="31dd9-162">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31dd9-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-163">Requirements</span></span>

|<span data-ttu-id="31dd9-164">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-164">Requirement</span></span>| <span data-ttu-id="31dd9-165">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-166">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-167">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-167">1.0</span></span>|
|[<span data-ttu-id="31dd9-168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-169">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-171">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="31dd9-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="31dd9-172">restUrl :String</span></span>

<span data-ttu-id="31dd9-173">Obtém o URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="31dd9-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="31dd9-174">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](https://docs.microsoft.com/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="31dd9-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="31dd9-175">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31dd9-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="31dd9-p105">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="31dd9-178">Tipo:</span><span class="sxs-lookup"><span data-stu-id="31dd9-178">Type:</span></span>

*   <span data-ttu-id="31dd9-179">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31dd9-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-180">Requirements</span></span>

|<span data-ttu-id="31dd9-181">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-181">Requirement</span></span>| <span data-ttu-id="31dd9-182">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-183">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-184">1.5</span><span class="sxs-lookup"><span data-stu-id="31dd9-184">1.5</span></span> |
|[<span data-ttu-id="31dd9-185">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-186">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-188">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="31dd9-189">Métodos</span><span class="sxs-lookup"><span data-stu-id="31dd9-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="31dd9-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="31dd9-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="31dd9-191">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="31dd9-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="31dd9-192">Atualmente, os tipos de evento compatíveis são `Office.EventType.ItemChanged` e `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-193">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-193">Parameters:</span></span>

| <span data-ttu-id="31dd9-194">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-194">Name</span></span> | <span data-ttu-id="31dd9-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-195">Type</span></span> | <span data-ttu-id="31dd9-196">Atributos</span><span class="sxs-lookup"><span data-stu-id="31dd9-196">Attributes</span></span> | <span data-ttu-id="31dd9-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="31dd9-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="31dd9-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="31dd9-199">O evento que deve chamar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="31dd9-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="31dd9-200">Função</span><span class="sxs-lookup"><span data-stu-id="31dd9-200">Function</span></span> || <span data-ttu-id="31dd9-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um literal de objeto. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="31dd9-204">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-204">Object</span></span> | <span data-ttu-id="31dd9-205">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-205">&lt;optional&gt;</span></span> | <span data-ttu-id="31dd9-206">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="31dd9-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="31dd9-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-207">Object</span></span> | <span data-ttu-id="31dd9-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-208">&lt;optional&gt;</span></span> | <span data-ttu-id="31dd9-209">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="31dd9-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="31dd9-210">função</span><span class="sxs-lookup"><span data-stu-id="31dd9-210">function</span></span>| <span data-ttu-id="31dd9-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-211">&lt;optional&gt;</span></span>|<span data-ttu-id="31dd9-212">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31dd9-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-213">Requirements</span></span>

|<span data-ttu-id="31dd9-214">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-214">Requirement</span></span>| <span data-ttu-id="31dd9-215">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-216">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-217">1.5</span><span class="sxs-lookup"><span data-stu-id="31dd9-217">1.5</span></span> |
|[<span data-ttu-id="31dd9-218">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-219">ReadItem</span></span> |
|[<span data-ttu-id="31dd9-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-221">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="31dd9-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="31dd9-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="31dd9-224">Converte uma ID de item formatada para REST em formato EWS.</span><span class="sxs-lookup"><span data-stu-id="31dd9-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-225">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-228">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-228">Parameters:</span></span>

|<span data-ttu-id="31dd9-229">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-229">Name</span></span>| <span data-ttu-id="31dd9-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-230">Type</span></span>| <span data-ttu-id="31dd9-231">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="31dd9-232">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-232">String</span></span>|<span data-ttu-id="31dd9-233">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="31dd9-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="31dd9-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="31dd9-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="31dd9-235">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="31dd9-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-236">Requirements</span></span>

|<span data-ttu-id="31dd9-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-237">Requirement</span></span>| <span data-ttu-id="31dd9-238">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-239">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-240">1.3</span><span class="sxs-lookup"><span data-stu-id="31dd9-240">1.3</span></span>|
|[<span data-ttu-id="31dd9-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-242">Restrito</span><span class="sxs-lookup"><span data-stu-id="31dd9-242">Restricted</span></span>|
|[<span data-ttu-id="31dd9-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-244">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31dd9-245">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31dd9-245">Returns:</span></span>

<span data-ttu-id="31dd9-246">Tipo: sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="31dd9-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="31dd9-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="31dd9-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="31dd9-249">Obtém um dicionário contendo informações da hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="31dd9-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="31dd9-p108">As datas e horas usadas por um aplicativo de email para o Outlook ou o aplicativo Web do Outlook podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o o aplicativo Web do Outlook usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="31dd9-p109">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no o aplicativo Web do Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-255">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-255">Parameters:</span></span>

|<span data-ttu-id="31dd9-256">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-256">Name</span></span>| <span data-ttu-id="31dd9-257">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-257">Type</span></span>| <span data-ttu-id="31dd9-258">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="31dd9-259">Data</span><span class="sxs-lookup"><span data-stu-id="31dd9-259">Date</span></span>|<span data-ttu-id="31dd9-260">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="31dd9-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-261">Requirements</span></span>

|<span data-ttu-id="31dd9-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-262">Requirement</span></span>| <span data-ttu-id="31dd9-263">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-264">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-265">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-265">1.0</span></span>|
|[<span data-ttu-id="31dd9-266">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-267">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-268">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-269">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31dd9-270">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31dd9-270">Returns:</span></span>

<span data-ttu-id="31dd9-271">Tipo: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="31dd9-271">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="31dd9-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="31dd9-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="31dd9-273">Converte uma ID de item formatada para EWS em formato REST.</span><span class="sxs-lookup"><span data-stu-id="31dd9-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-274">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-p110">As IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-277">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-277">Parameters:</span></span>

|<span data-ttu-id="31dd9-278">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-278">Name</span></span>| <span data-ttu-id="31dd9-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-279">Type</span></span>| <span data-ttu-id="31dd9-280">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="31dd9-281">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-281">String</span></span>|<span data-ttu-id="31dd9-282">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="31dd9-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="31dd9-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="31dd9-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="31dd9-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="31dd9-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-285">Requirements</span></span>

|<span data-ttu-id="31dd9-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-286">Requirement</span></span>| <span data-ttu-id="31dd9-287">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-288">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-289">1.3</span><span class="sxs-lookup"><span data-stu-id="31dd9-289">1.3</span></span>|
|[<span data-ttu-id="31dd9-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-291">Restrito</span><span class="sxs-lookup"><span data-stu-id="31dd9-291">Restricted</span></span>|
|[<span data-ttu-id="31dd9-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-293">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31dd9-294">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31dd9-294">Returns:</span></span>

<span data-ttu-id="31dd9-295">Tipo: sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="31dd9-296">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="31dd9-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="31dd9-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="31dd9-298">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="31dd9-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="31dd9-299">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="31dd9-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-300">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-300">Parameters:</span></span>

|<span data-ttu-id="31dd9-301">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-301">Name</span></span>| <span data-ttu-id="31dd9-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-302">Type</span></span>| <span data-ttu-id="31dd9-303">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="31dd9-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="31dd9-304">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="31dd9-305">O valor temporal local a converter.</span><span class="sxs-lookup"><span data-stu-id="31dd9-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-306">Requirements</span></span>

|<span data-ttu-id="31dd9-307">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-307">Requirement</span></span>| <span data-ttu-id="31dd9-308">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-309">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-310">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-310">1.0</span></span>|
|[<span data-ttu-id="31dd9-311">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-312">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-314">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="31dd9-315">Retorna:</span><span class="sxs-lookup"><span data-stu-id="31dd9-315">Returns:</span></span>

<span data-ttu-id="31dd9-316">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="31dd9-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="31dd9-317">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="31dd9-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="31dd9-318">Data</span><span class="sxs-lookup"><span data-stu-id="31dd9-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="31dd9-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="31dd9-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="31dd9-320">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="31dd9-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-321">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-322">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="31dd9-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="31dd9-p111">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="31dd9-325">No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31dd9-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="31dd9-326">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="31dd9-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-327">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-327">Parameters:</span></span>

|<span data-ttu-id="31dd9-328">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-328">Name</span></span>| <span data-ttu-id="31dd9-329">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-329">Type</span></span>| <span data-ttu-id="31dd9-330">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="31dd9-331">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-331">String</span></span>|<span data-ttu-id="31dd9-332">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-333">Requirements</span></span>

|<span data-ttu-id="31dd9-334">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-334">Requirement</span></span>| <span data-ttu-id="31dd9-335">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-336">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-337">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-337">1.0</span></span>|
|[<span data-ttu-id="31dd9-338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-339">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-341">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-342">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="31dd9-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="31dd9-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="31dd9-344">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="31dd9-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-345">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-346">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="31dd9-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="31dd9-347">No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31dd9-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="31dd9-348">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="31dd9-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="31dd9-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-351">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-351">Parameters:</span></span>

|<span data-ttu-id="31dd9-352">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-352">Name</span></span>| <span data-ttu-id="31dd9-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-353">Type</span></span>| <span data-ttu-id="31dd9-354">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="31dd9-355">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-355">String</span></span>|<span data-ttu-id="31dd9-356">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="31dd9-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-357">Requirements</span></span>

|<span data-ttu-id="31dd9-358">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-358">Requirement</span></span>| <span data-ttu-id="31dd9-359">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-360">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-361">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-361">1.0</span></span>|
|[<span data-ttu-id="31dd9-362">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-363">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-364">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-365">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-366">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="31dd9-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="31dd9-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="31dd9-368">Exibe um formulário para criar um novo compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="31dd9-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-369">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="31dd9-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="31dd9-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="31dd9-p114">No aplicativo Web do Outlook e no OWA para Dispositivos, esse método sempre exibe um formulário com um campo de participantes. Se você não especificar nenhum participante como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="31dd9-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees` ou `resources`, esse método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="31dd9-377">Se qualquer dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="31dd9-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-378">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-379">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="31dd9-379">Note: All parameters are optional.</span></span>

|<span data-ttu-id="31dd9-380">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-380">Name</span></span>| <span data-ttu-id="31dd9-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-381">Type</span></span>| <span data-ttu-id="31dd9-382">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="31dd9-383">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-383">Object</span></span> | <span data-ttu-id="31dd9-384">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="31dd9-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="31dd9-385">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="31dd9-p116">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="31dd9-388">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="31dd9-p117">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="31dd9-391">Data</span><span class="sxs-lookup"><span data-stu-id="31dd9-391">Date</span></span> | <span data-ttu-id="31dd9-392">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="31dd9-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="31dd9-393">Data</span><span class="sxs-lookup"><span data-stu-id="31dd9-393">Date</span></span> | <span data-ttu-id="31dd9-394">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="31dd9-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="31dd9-395">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-395">String</span></span> | <span data-ttu-id="31dd9-p118">Uma sequência de caracteres que contém o local do compromisso. Ela está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="31dd9-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="31dd9-p119">Uma matriz de sequências de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="31dd9-401">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-401">String</span></span> | <span data-ttu-id="31dd9-p120">Uma sequência de caracteres que contém o assunto do compromisso. Ela está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="31dd9-404">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-404">String</span></span> | <span data-ttu-id="31dd9-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31dd9-407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-407">Requirements</span></span>

|<span data-ttu-id="31dd9-408">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-408">Requirement</span></span>| <span data-ttu-id="31dd9-409">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-410">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-410">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-411">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-411">1.0</span></span>|
|[<span data-ttu-id="31dd9-412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-413">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-414">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-415">Leitura</span><span class="sxs-lookup"><span data-stu-id="31dd9-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="31dd9-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="31dd9-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="31dd9-418">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="31dd9-419">O método `displayNewMessageForm` abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="31dd9-420">Quando os parâmetros são especificados, os campos do formulário de mensagem são preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="31dd9-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="31dd9-421">Se qualquer dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="31dd9-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-422">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-423">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="31dd9-423">Note: All parameters are optional.</span></span>

|<span data-ttu-id="31dd9-424">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-424">Name</span></span>| <span data-ttu-id="31dd9-425">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-425">Type</span></span>| <span data-ttu-id="31dd9-426">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="31dd9-427">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-427">Object</span></span> | <span data-ttu-id="31dd9-428">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="31dd9-429">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="31dd9-430">Uma matriz de sequência de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Para.</span><span class="sxs-lookup"><span data-stu-id="31dd9-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="31dd9-431">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="31dd9-432">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="31dd9-433">Uma matriz de sequência de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cc.</span><span class="sxs-lookup"><span data-stu-id="31dd9-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="31dd9-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="31dd9-435">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="31dd9-436">Uma matriz de sequência de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="31dd9-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="31dd9-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="31dd9-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="31dd9-438">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-438">String</span></span> | <span data-ttu-id="31dd9-439">Uma sequência de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-439">A string containing the subject of the message.</span></span> <span data-ttu-id="31dd9-440">A sequência de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31dd9-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="31dd9-441">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-441">String</span></span> | <span data-ttu-id="31dd9-442">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-442">The HTML body of the message.</span></span> <span data-ttu-id="31dd9-443">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="31dd9-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="31dd9-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="31dd9-445">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="31dd9-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="31dd9-446">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-446">String</span></span> | <span data-ttu-id="31dd9-p128">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="31dd9-449">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-449">String</span></span> | <span data-ttu-id="31dd9-450">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="31dd9-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="31dd9-451">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-451">String</span></span> | <span data-ttu-id="31dd9-p129">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="31dd9-454">Booleano</span><span class="sxs-lookup"><span data-stu-id="31dd9-454">Boolean</span></span> | <span data-ttu-id="31dd9-p130">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado em linha no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="31dd9-457">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-457">String</span></span> | <span data-ttu-id="31dd9-458">Usado somente se `type` estiver definido para `item`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="31dd9-459">A ID do item do EWS do email existente que deseja anexar na nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="31dd9-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="31dd9-460">Isso é uma sequência de caracteres com até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="31dd9-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="31dd9-461">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-461">Requirements</span></span>

|<span data-ttu-id="31dd9-462">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-462">Requirement</span></span>| <span data-ttu-id="31dd9-463">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-464">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-464">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-465">1.6</span><span class="sxs-lookup"><span data-stu-id="31dd9-465">-16</span></span> |
|[<span data-ttu-id="31dd9-466">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-467">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-468">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-469">Leitura</span><span class="sxs-lookup"><span data-stu-id="31dd9-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-470">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-470">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="31dd9-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="31dd9-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="31dd9-472">Obtém uma sequência de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="31dd9-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="31dd9-p132">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-475">É recomendável que suplementos usem as APIs REST em vez de Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="31dd9-475">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="31dd9-476">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="31dd9-476">**REST Tokens**</span></span>

<span data-ttu-id="31dd9-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="31dd9-480">O suplemento deve usar a propriedade `restUrl` para determinar o URL correto a ser usado ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="31dd9-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="31dd9-481">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="31dd9-481">**EWS Tokens**</span></span>

<span data-ttu-id="31dd9-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="31dd9-484">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="31dd9-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-485">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-485">Parameters:</span></span>

|<span data-ttu-id="31dd9-486">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-486">Name</span></span>| <span data-ttu-id="31dd9-487">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-487">Type</span></span>| <span data-ttu-id="31dd9-488">Atributos</span><span class="sxs-lookup"><span data-stu-id="31dd9-488">Attributes</span></span>| <span data-ttu-id="31dd9-489">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="31dd9-490">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-490">Object</span></span> | <span data-ttu-id="31dd9-491">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-491">&lt;optional&gt;</span></span> | <span data-ttu-id="31dd9-492">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="31dd9-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="31dd9-493">Booleano</span><span class="sxs-lookup"><span data-stu-id="31dd9-493">Boolean</span></span> |  <span data-ttu-id="31dd9-494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-494">&lt;optional&gt;</span></span> | <span data-ttu-id="31dd9-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="31dd9-497">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-497">Object</span></span> |  <span data-ttu-id="31dd9-498">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-498">&lt;optional&gt;</span></span> | <span data-ttu-id="31dd9-499">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="31dd9-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="31dd9-500">função</span><span class="sxs-lookup"><span data-stu-id="31dd9-500">function</span></span>||<span data-ttu-id="31dd9-p136">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-503">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-503">Requirements</span></span>

|<span data-ttu-id="31dd9-504">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-504">Requirement</span></span>| <span data-ttu-id="31dd9-505">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-506">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-507">1.5</span><span class="sxs-lookup"><span data-stu-id="31dd9-507">1.5</span></span> |
|[<span data-ttu-id="31dd9-508">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-509">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-510">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-511">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-512">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-512">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="31dd9-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="31dd9-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="31dd9-514">Obtém uma sequência de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="31dd9-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="31dd9-p137">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="31dd9-p138">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="31dd9-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="31dd9-520">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="31dd9-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="31dd9-p139">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-523">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-523">Parameters:</span></span>

|<span data-ttu-id="31dd9-524">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-524">Name</span></span>| <span data-ttu-id="31dd9-525">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-525">Type</span></span>| <span data-ttu-id="31dd9-526">Atributos</span><span class="sxs-lookup"><span data-stu-id="31dd9-526">Attributes</span></span>| <span data-ttu-id="31dd9-527">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="31dd9-528">função</span><span class="sxs-lookup"><span data-stu-id="31dd9-528">function</span></span>||<span data-ttu-id="31dd9-p140">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="31dd9-531">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-531">Object</span></span>| <span data-ttu-id="31dd9-532">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-532">&lt;optional&gt;</span></span>|<span data-ttu-id="31dd9-533">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="31dd9-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-534">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-534">Requirements</span></span>

|<span data-ttu-id="31dd9-535">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-535">Requirement</span></span>| <span data-ttu-id="31dd9-536">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-537">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-537">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-538">1.3</span><span class="sxs-lookup"><span data-stu-id="31dd9-538">1.3</span></span>|
|[<span data-ttu-id="31dd9-539">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-540">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-541">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-542">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-543">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="31dd9-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="31dd9-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="31dd9-545">Obtém um token que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="31dd9-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="31dd9-546">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="31dd9-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-547">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-547">Parameters:</span></span>

|<span data-ttu-id="31dd9-548">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-548">Name</span></span>| <span data-ttu-id="31dd9-549">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-549">Type</span></span>| <span data-ttu-id="31dd9-550">Atributos</span><span class="sxs-lookup"><span data-stu-id="31dd9-550">Attributes</span></span>| <span data-ttu-id="31dd9-551">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="31dd9-552">função</span><span class="sxs-lookup"><span data-stu-id="31dd9-552">function</span></span>||<span data-ttu-id="31dd9-553">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31dd9-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="31dd9-554">O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="31dd9-555">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-555">Object</span></span>| <span data-ttu-id="31dd9-556">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-556">&lt;optional&gt;</span></span>|<span data-ttu-id="31dd9-557">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="31dd9-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-558">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-558">Requirements</span></span>

|<span data-ttu-id="31dd9-559">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-559">Requirement</span></span>| <span data-ttu-id="31dd9-560">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-561">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-561">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-562">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-562">1.0</span></span>|
|[<span data-ttu-id="31dd9-563">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31dd9-564">ReadItem</span></span>|
|[<span data-ttu-id="31dd9-565">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-566">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-567">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="31dd9-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="31dd9-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="31dd9-569">Faz uma solicitação assíncrona em um serviço dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="31dd9-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-570">Esse método não pode ser usado nos seguintes cenários.</span><span class="sxs-lookup"><span data-stu-id="31dd9-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="31dd9-571">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="31dd9-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="31dd9-572">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="31dd9-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="31dd9-573">Nesses casos, os suplementos devem [usar APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="31dd9-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="31dd9-574">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="31dd9-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="31dd9-575">Para obter uma lista de operações EWS compatíveis, consulte [Chamar serviços Web de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="31dd9-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="31dd9-576">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="31dd9-577">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="31dd9-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="31dd9-p142">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, consulte [Especificar permissões para acesso do suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="31dd9-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="31dd9-580">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o servidor de Acesso para Cliente, para que o método `makeEwsRequestAsync` possa realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="31dd9-580">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="31dd9-581">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="31dd9-581">Version differences</span></span>

<span data-ttu-id="31dd9-582">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="31dd9-p143">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Para determinar qual versão do Outlook está em execução, use a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="31dd9-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="31dd9-586">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="31dd9-586">Parameters:</span></span>

|<span data-ttu-id="31dd9-587">Nome</span><span class="sxs-lookup"><span data-stu-id="31dd9-587">Name</span></span>| <span data-ttu-id="31dd9-588">Tipo</span><span class="sxs-lookup"><span data-stu-id="31dd9-588">Type</span></span>| <span data-ttu-id="31dd9-589">Atributos</span><span class="sxs-lookup"><span data-stu-id="31dd9-589">Attributes</span></span>| <span data-ttu-id="31dd9-590">Descrição</span><span class="sxs-lookup"><span data-stu-id="31dd9-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="31dd9-591">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="31dd9-591">String</span></span>||<span data-ttu-id="31dd9-592">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="31dd9-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="31dd9-593">função</span><span class="sxs-lookup"><span data-stu-id="31dd9-593">function</span></span>||<span data-ttu-id="31dd9-594">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="31dd9-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="31dd9-595">O resultado XML da chamada do EWS é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="31dd9-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="31dd9-596">Se o tamanho do resultado exceder 1 MB, uma mensagem de erro será exibida em vez disso.</span><span class="sxs-lookup"><span data-stu-id="31dd9-596">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="31dd9-597">Objeto</span><span class="sxs-lookup"><span data-stu-id="31dd9-597">Object</span></span>| <span data-ttu-id="31dd9-598">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="31dd9-598">&lt;optional&gt;</span></span>|<span data-ttu-id="31dd9-599">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="31dd9-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31dd9-600">Requisitos</span><span class="sxs-lookup"><span data-stu-id="31dd9-600">Requirements</span></span>

|<span data-ttu-id="31dd9-601">Requisito</span><span class="sxs-lookup"><span data-stu-id="31dd9-601">Requirement</span></span>| <span data-ttu-id="31dd9-602">Valor</span><span class="sxs-lookup"><span data-stu-id="31dd9-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="31dd9-603">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="31dd9-603">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31dd9-604">1.0</span><span class="sxs-lookup"><span data-stu-id="31dd9-604">1.0</span></span>|
|[<span data-ttu-id="31dd9-605">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="31dd9-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31dd9-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="31dd9-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="31dd9-607">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="31dd9-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31dd9-608">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="31dd9-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="31dd9-609">Exemplo</span><span class="sxs-lookup"><span data-stu-id="31dd9-609">Example</span></span>

<span data-ttu-id="31dd9-610">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="31dd9-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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