
# <a name="mailbox"></a><span data-ttu-id="6f0cb-101">caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-101">mailbox</span></span>

### <span data-ttu-id="6f0cb-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="6f0cb-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6f0cb-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-105">Requirements</span></span>

|<span data-ttu-id="6f0cb-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-106">Requirement</span></span>| <span data-ttu-id="6f0cb-107">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-109">1.0</span></span>|
|[<span data-ttu-id="6f0cb-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-111">Restricted</span></span>|
|[<span data-ttu-id="6f0cb-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-113">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="6f0cb-114">Namespaces</span><span class="sxs-lookup"><span data-stu-id="6f0cb-114">Namespaces</span></span>

<span data-ttu-id="6f0cb-115">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="6f0cb-116">[item](Office.context.mailbox.item.md): Fornece métodos e propriedades para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="6f0cb-117">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="6f0cb-118">Membros</span><span class="sxs-lookup"><span data-stu-id="6f0cb-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="6f0cb-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-119">ewsUrl :String</span></span>

<span data-ttu-id="6f0cb-p102">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-122">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-122">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-p103">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="6f0cb-125">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="6f0cb-p104">No modo redigir, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="6f0cb-128">Tipo:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-128">Type:</span></span>

*   <span data-ttu-id="6f0cb-129">String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6f0cb-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-130">Requirements</span></span>

|<span data-ttu-id="6f0cb-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-131">Requirement</span></span>| <span data-ttu-id="6f0cb-132">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-133">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-134">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-134">1.0</span></span>|
|[<span data-ttu-id="6f0cb-135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-136">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-138">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-138">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="6f0cb-139">Métodos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="6f0cb-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="6f0cb-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="6f0cb-141">Converte uma ID de item formatada para REST em formato EWS.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-142">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-142">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-p105">As IDs de itens recuperadas por meio de uma API REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada em REST no formato adequado para o EWS.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-145">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-145">Parameters:</span></span>

|<span data-ttu-id="6f0cb-146">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-146">Name</span></span>| <span data-ttu-id="6f0cb-147">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-147">Type</span></span>| <span data-ttu-id="6f0cb-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6f0cb-149">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6f0cb-149">String</span></span>|<span data-ttu-id="6f0cb-150">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="6f0cb-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="6f0cb-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="6f0cb-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="6f0cb-152">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-153">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-153">Requirements</span></span>

|<span data-ttu-id="6f0cb-154">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-154">Requirement</span></span>| <span data-ttu-id="6f0cb-155">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-156">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-157">1.3</span><span class="sxs-lookup"><span data-stu-id="6f0cb-157">1.3</span></span>|
|[<span data-ttu-id="6f0cb-158">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-159">Restrito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-159">Restricted</span></span>|
|[<span data-ttu-id="6f0cb-160">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-161">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-161">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6f0cb-162">Retorna:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-162">Returns:</span></span>

<span data-ttu-id="6f0cb-163">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="6f0cb-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-164">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="6f0cb-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="6f0cb-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="6f0cb-166">Obtém um dicionário contendo informações da hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="6f0cb-p106">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; O Outlook Web App usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve manipular valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário esperado pelo usuário.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p106">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="6f0cb-p107">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p107">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-172">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-172">Parameters:</span></span>

|<span data-ttu-id="6f0cb-173">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-173">Name</span></span>| <span data-ttu-id="6f0cb-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-174">Type</span></span>| <span data-ttu-id="6f0cb-175">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="6f0cb-176">Date</span><span class="sxs-lookup"><span data-stu-id="6f0cb-176">Date</span></span>|<span data-ttu-id="6f0cb-177">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="6f0cb-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-178">Requirements</span></span>

|<span data-ttu-id="6f0cb-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-179">Requirement</span></span>| <span data-ttu-id="6f0cb-180">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-181">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-182">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-182">1.0</span></span>|
|[<span data-ttu-id="6f0cb-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-184">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-186">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-186">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6f0cb-187">Retorna:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-187">Returns:</span></span>

<span data-ttu-id="6f0cb-188">Tipo: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="6f0cb-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="6f0cb-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="6f0cb-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="6f0cb-190">Converte uma ID de item formatada para EWS em formato REST.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-191">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-191">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-p108">As IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API de email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](http://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-194">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-194">Parameters:</span></span>

|<span data-ttu-id="6f0cb-195">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-195">Name</span></span>| <span data-ttu-id="6f0cb-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-196">Type</span></span>| <span data-ttu-id="6f0cb-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6f0cb-198">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6f0cb-198">String</span></span>|<span data-ttu-id="6f0cb-199">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="6f0cb-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="6f0cb-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="6f0cb-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="6f0cb-201">Um valor indicando a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-202">Requirements</span></span>

|<span data-ttu-id="6f0cb-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-203">Requirement</span></span>| <span data-ttu-id="6f0cb-204">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-205">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-206">1.3</span><span class="sxs-lookup"><span data-stu-id="6f0cb-206">1.3</span></span>|
|[<span data-ttu-id="6f0cb-207">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-208">Restrito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-208">Restricted</span></span>|
|[<span data-ttu-id="6f0cb-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-210">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-210">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6f0cb-211">Retorna:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-211">Returns:</span></span>

<span data-ttu-id="6f0cb-212">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="6f0cb-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-213">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="6f0cb-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="6f0cb-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="6f0cb-215">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="6f0cb-216">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-217">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-217">Parameters:</span></span>

|<span data-ttu-id="6f0cb-218">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-218">Name</span></span>| <span data-ttu-id="6f0cb-219">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-219">Type</span></span>| <span data-ttu-id="6f0cb-220">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="6f0cb-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="6f0cb-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="6f0cb-222">O valor de hora local a ser convertido.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-223">Requirements</span></span>

|<span data-ttu-id="6f0cb-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-224">Requirement</span></span>| <span data-ttu-id="6f0cb-225">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-226">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-227">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-227">1.0</span></span>|
|[<span data-ttu-id="6f0cb-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-229">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-230">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-231">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-231">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6f0cb-232">Retorna:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-232">Returns:</span></span>

<span data-ttu-id="6f0cb-233">Um objeto Date com o horário expresso em UTC.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="6f0cb-234">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="6f0cb-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6f0cb-235">Date</span><span class="sxs-lookup"><span data-stu-id="6f0cb-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="6f0cb-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="6f0cb-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="6f0cb-237">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-238">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-238">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-239">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="6f0cb-p109">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso principal de uma série recorrente, mas você não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p109">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="6f0cb-242">No aplicativo Web do Outlook, esse método abre o formulário especificado somente se o corpo do formulário for menor ou igual a um número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="6f0cb-243">Se o identificador de item especificado não identificar um compromisso existente, um painel em branco será aberto no computador ou no dispositivo cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-244">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-244">Parameters:</span></span>

|<span data-ttu-id="6f0cb-245">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-245">Name</span></span>| <span data-ttu-id="6f0cb-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-246">Type</span></span>| <span data-ttu-id="6f0cb-247">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6f0cb-248">String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-248">String</span></span>|<span data-ttu-id="6f0cb-249">O identificador de serviços da Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-250">Requirements</span></span>

|<span data-ttu-id="6f0cb-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-251">Requirement</span></span>| <span data-ttu-id="6f0cb-252">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-253">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-254">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-254">1.0</span></span>|
|[<span data-ttu-id="6f0cb-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-256">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-258">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-258">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-259">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="6f0cb-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="6f0cb-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="6f0cb-261">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-262">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-262">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-263">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="6f0cb-264">No aplicativo Web do Outlook, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual a um número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="6f0cb-265">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida uma mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="6f0cb-p110">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-268">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-268">Parameters:</span></span>

|<span data-ttu-id="6f0cb-269">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-269">Name</span></span>| <span data-ttu-id="6f0cb-270">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-270">Type</span></span>| <span data-ttu-id="6f0cb-271">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6f0cb-272">String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-272">String</span></span>|<span data-ttu-id="6f0cb-273">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-274">Requirements</span></span>

|<span data-ttu-id="6f0cb-275">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-275">Requirement</span></span>| <span data-ttu-id="6f0cb-276">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-277">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-278">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-278">1.0</span></span>|
|[<span data-ttu-id="6f0cb-279">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-280">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-281">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-282">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-282">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-283">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-283">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="6f0cb-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="6f0cb-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="6f0cb-285">Exibe um formulário para criar um novo compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-286">Esse método não  é suportado no Outlook para iOS nem no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-286">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6f0cb-p111">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="6f0cb-p112">No aplicativo Web do Outlook e no OWA para Dispositivos, esse método sempre exibe um formulário com um campo de participantes. Se você não especificar nenhum participante como argumento de entrada, o método exibirá um formulário com um botão **Salvar** . Se você especificar participantes, o formulário incluirá os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p112">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="6f0cb-p113">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees` ou `resources`, o método exibirá um formulário de reunião com um botão **Enviar** . Se você não especificar destinatários, o método exibirá um formulário de compromisso com um botão **Salvar & Fechar**.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="6f0cb-294">Se algum dos parâmetros exceder os limites de tamanho especificados ou se um nome de parâmetro desconhecido for especificado, será gerada uma exceção.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-295">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-295">Parameters:</span></span>

|<span data-ttu-id="6f0cb-296">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-296">Name</span></span>| <span data-ttu-id="6f0cb-297">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-297">Type</span></span>| <span data-ttu-id="6f0cb-298">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="6f0cb-299">Objeto</span><span class="sxs-lookup"><span data-stu-id="6f0cb-299">Object</span></span> | <span data-ttu-id="6f0cb-300">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="6f0cb-301">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="6f0cb-p114">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="6f0cb-304">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="6f0cb-p115">Uma matriz de sequências de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="6f0cb-307">Date</span><span class="sxs-lookup"><span data-stu-id="6f0cb-307">Date</span></span> | <span data-ttu-id="6f0cb-308">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="6f0cb-309">Date</span><span class="sxs-lookup"><span data-stu-id="6f0cb-309">Date</span></span> | <span data-ttu-id="6f0cb-310">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="6f0cb-311">String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-311">String</span></span> | <span data-ttu-id="6f0cb-p116">Uma sequência de caracteres que contém o local do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="6f0cb-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="6f0cb-p117">Uma matriz de sequências de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="6f0cb-317">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6f0cb-317">String</span></span> | <span data-ttu-id="6f0cb-p118">Uma sequência de caracteres que contém o assunto do compromisso. Está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="6f0cb-320">String</span><span class="sxs-lookup"><span data-stu-id="6f0cb-320">String</span></span> | <span data-ttu-id="6f0cb-p119">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6f0cb-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-323">Requirements</span></span>

|<span data-ttu-id="6f0cb-324">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-324">Requirement</span></span>| <span data-ttu-id="6f0cb-325">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-326">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-327">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-327">1.0</span></span>|
|[<span data-ttu-id="6f0cb-328">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-329">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-330">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-331">Leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-332">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="6f0cb-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6f0cb-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="6f0cb-334">Obtém uma sequência de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="6f0cb-p120">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p120">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="6f0cb-p121">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p121">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="6f0cb-340">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="6f0cb-p122">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p122">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-343">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-343">Parameters:</span></span>

|<span data-ttu-id="6f0cb-344">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-344">Name</span></span>| <span data-ttu-id="6f0cb-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-345">Type</span></span>| <span data-ttu-id="6f0cb-346">Atributos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-346">Attributes</span></span>| <span data-ttu-id="6f0cb-347">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6f0cb-348">function</span><span class="sxs-lookup"><span data-stu-id="6f0cb-348">function</span></span>||<span data-ttu-id="6f0cb-p123">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p123">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="6f0cb-351">Objeto</span><span class="sxs-lookup"><span data-stu-id="6f0cb-351">Object</span></span>| <span data-ttu-id="6f0cb-352">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-352">&lt;optional&gt;</span></span>|<span data-ttu-id="6f0cb-353">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-354">Requirements</span></span>

|<span data-ttu-id="6f0cb-355">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-355">Requirement</span></span>| <span data-ttu-id="6f0cb-356">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-357">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-358">1.3</span><span class="sxs-lookup"><span data-stu-id="6f0cb-358">1.3</span></span>|
|[<span data-ttu-id="6f0cb-359">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-360">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-361">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-362">Redigir e ler</span><span class="sxs-lookup"><span data-stu-id="6f0cb-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-363">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-363">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="6f0cb-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6f0cb-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="6f0cb-365">Obtém um token que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="6f0cb-366">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-367">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-367">Parameters:</span></span>

|<span data-ttu-id="6f0cb-368">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-368">Name</span></span>| <span data-ttu-id="6f0cb-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-369">Type</span></span>| <span data-ttu-id="6f0cb-370">Atributos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-370">Attributes</span></span>| <span data-ttu-id="6f0cb-371">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6f0cb-372">function</span><span class="sxs-lookup"><span data-stu-id="6f0cb-372">function</span></span>||<span data-ttu-id="6f0cb-373">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6f0cb-374">O token é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="6f0cb-375">Objeto</span><span class="sxs-lookup"><span data-stu-id="6f0cb-375">Object</span></span>| <span data-ttu-id="6f0cb-376">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-376">&lt;optional&gt;</span></span>|<span data-ttu-id="6f0cb-377">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-378">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-378">Requirements</span></span>

|<span data-ttu-id="6f0cb-379">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-379">Requirement</span></span>| <span data-ttu-id="6f0cb-380">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-381">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-382">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-382">1.0</span></span>|
|[<span data-ttu-id="6f0cb-383">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6f0cb-384">ReadItem</span></span>|
|[<span data-ttu-id="6f0cb-385">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-386">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-386">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-387">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-387">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="6f0cb-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6f0cb-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="6f0cb-389">Faz uma solicitação assíncrona em um dos Serviços Web do Exchange (EWS) no Exchange Server que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-390">Esse método não é suportado nos seguintes cenários.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="6f0cb-391">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="6f0cb-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="6f0cb-392">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="6f0cb-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="6f0cb-393">Nesses casos, os suplementos devem [usar APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="6f0cb-394">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="6f0cb-395">Para obter uma lista de operações EWS compatíveis, confira [Chamar serviços Web de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="6f0cb-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="6f0cb-396">Não é possível solicitar os itens associados à pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="6f0cb-397">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-397">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="6f0cb-p125">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para obter mais informações sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso do suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="6f0cb-400">O administrador do servidor deve definir `OAuthAuthentication` como verdadeiro no diretório EWS do Servidor de Acesso para Cliente para ativar o método `makeEwsRequestAsync` para fazer solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-400">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="6f0cb-401">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="6f0cb-401">Version differences</span></span>

<span data-ttu-id="6f0cb-402">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="6f0cb-p126">Você não precisa definir o valor de codificação quando seu aplicativo de email estiver sendo executado no Outlook na Web. Você pode determinar se o seu aplicativo de email está sendo executado no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar qual versão do Outlook está sendo executada usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6f0cb-406">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="6f0cb-406">Parameters:</span></span>

|<span data-ttu-id="6f0cb-407">Nome</span><span class="sxs-lookup"><span data-stu-id="6f0cb-407">Name</span></span>| <span data-ttu-id="6f0cb-408">Tipo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-408">Type</span></span>| <span data-ttu-id="6f0cb-409">Atributos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-409">Attributes</span></span>| <span data-ttu-id="6f0cb-410">Descrição</span><span class="sxs-lookup"><span data-stu-id="6f0cb-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6f0cb-411">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="6f0cb-411">String</span></span>||<span data-ttu-id="6f0cb-412">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="6f0cb-413">function</span><span class="sxs-lookup"><span data-stu-id="6f0cb-413">function</span></span>||<span data-ttu-id="6f0cb-414">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6f0cb-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6f0cb-415">O resultado XML da chamada do EWS é fornecido como uma sequência de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="6f0cb-416">Se o resultado exceder 1 MB, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-416">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="6f0cb-417">Objeto</span><span class="sxs-lookup"><span data-stu-id="6f0cb-417">Object</span></span>| <span data-ttu-id="6f0cb-418">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="6f0cb-418">&lt;optional&gt;</span></span>|<span data-ttu-id="6f0cb-419">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f0cb-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6f0cb-420">Requirements</span></span>

|<span data-ttu-id="6f0cb-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="6f0cb-421">Requirement</span></span>| <span data-ttu-id="6f0cb-422">Valor</span><span class="sxs-lookup"><span data-stu-id="6f0cb-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f0cb-423">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6f0cb-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6f0cb-424">1.0</span><span class="sxs-lookup"><span data-stu-id="6f0cb-424">1.0</span></span>|
|[<span data-ttu-id="6f0cb-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6f0cb-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="6f0cb-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="6f0cb-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6f0cb-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6f0cb-428">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6f0cb-428">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6f0cb-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6f0cb-429">Example</span></span>

<span data-ttu-id="6f0cb-430">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="6f0cb-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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