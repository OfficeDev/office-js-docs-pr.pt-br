---
title: 'Office.context.mailbox: conjunto de requisitos da versão 1.5'
description: ''
ms.date: 08/06/2019
localization_priority: Priority
ms.openlocfilehash: 68912520250789b2259d59fb14387f97b24c5c7d
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268639"
---
# <a name="mailbox"></a><span data-ttu-id="af631-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="af631-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="af631-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="af631-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="af631-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="af631-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="af631-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-105">Requirements</span></span>

|<span data-ttu-id="af631-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-106">Requirement</span></span>| <span data-ttu-id="af631-107">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-109">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-109">1.0</span></span>|
|[<span data-ttu-id="af631-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="af631-111">Restricted</span></span>|
|[<span data-ttu-id="af631-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="af631-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="af631-114">Members and methods</span></span>

| <span data-ttu-id="af631-115">Membro</span><span class="sxs-lookup"><span data-stu-id="af631-115">Member</span></span> | <span data-ttu-id="af631-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="af631-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="af631-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="af631-118">Membro</span><span class="sxs-lookup"><span data-stu-id="af631-118">Member</span></span> |
| [<span data-ttu-id="af631-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="af631-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="af631-120">Membro</span><span class="sxs-lookup"><span data-stu-id="af631-120">Member</span></span> |
| [<span data-ttu-id="af631-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="af631-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="af631-122">Método</span><span class="sxs-lookup"><span data-stu-id="af631-122">Method</span></span> |
| [<span data-ttu-id="af631-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="af631-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="af631-124">Método</span><span class="sxs-lookup"><span data-stu-id="af631-124">Method</span></span> |
| [<span data-ttu-id="af631-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="af631-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="af631-126">Método</span><span class="sxs-lookup"><span data-stu-id="af631-126">Method</span></span> |
| [<span data-ttu-id="af631-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="af631-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="af631-128">Método</span><span class="sxs-lookup"><span data-stu-id="af631-128">Method</span></span> |
| [<span data-ttu-id="af631-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="af631-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="af631-130">Método</span><span class="sxs-lookup"><span data-stu-id="af631-130">Method</span></span> |
| [<span data-ttu-id="af631-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="af631-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="af631-132">Método</span><span class="sxs-lookup"><span data-stu-id="af631-132">Method</span></span> |
| [<span data-ttu-id="af631-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="af631-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="af631-134">Método</span><span class="sxs-lookup"><span data-stu-id="af631-134">Method</span></span> |
| [<span data-ttu-id="af631-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="af631-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="af631-136">Método</span><span class="sxs-lookup"><span data-stu-id="af631-136">Method</span></span> |
| [<span data-ttu-id="af631-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="af631-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="af631-138">Método</span><span class="sxs-lookup"><span data-stu-id="af631-138">Method</span></span> |
| [<span data-ttu-id="af631-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="af631-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="af631-140">Método</span><span class="sxs-lookup"><span data-stu-id="af631-140">Method</span></span> |
| [<span data-ttu-id="af631-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="af631-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="af631-142">Método</span><span class="sxs-lookup"><span data-stu-id="af631-142">Method</span></span> |
| [<span data-ttu-id="af631-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="af631-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="af631-144">Método</span><span class="sxs-lookup"><span data-stu-id="af631-144">Method</span></span> |
| [<span data-ttu-id="af631-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="af631-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="af631-146">Método</span><span class="sxs-lookup"><span data-stu-id="af631-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="af631-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="af631-147">Namespaces</span></span>

<span data-ttu-id="af631-148">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="af631-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="af631-149">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="af631-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="af631-150">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="af631-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="af631-151">Members</span><span class="sxs-lookup"><span data-stu-id="af631-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="af631-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="af631-152">ewsUrl :String</span></span>

<span data-ttu-id="af631-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="af631-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-155">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="af631-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="af631-158">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af631-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="af631-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="af631-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="af631-161">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-161">Type</span></span>

*   <span data-ttu-id="af631-162">String</span><span class="sxs-lookup"><span data-stu-id="af631-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af631-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-163">Requirements</span></span>

|<span data-ttu-id="af631-164">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-164">Requirement</span></span>| <span data-ttu-id="af631-165">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-166">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-167">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-167">1.0</span></span>|
|[<span data-ttu-id="af631-168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-169">ReadItem</span></span>|
|[<span data-ttu-id="af631-170">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="af631-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="af631-172">restUrl :String</span></span>

<span data-ttu-id="af631-173">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="af631-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="af631-174">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="af631-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="af631-175">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af631-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="af631-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="af631-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-178">Os clientes do Outlook conectados a instalações locais do Exchange 2016 ou posterior com um REST personalizado da URL configurada retornarão um valor inválido para `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="af631-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="af631-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-179">Type</span></span>

*   <span data-ttu-id="af631-180">String</span><span class="sxs-lookup"><span data-stu-id="af631-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af631-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-181">Requirements</span></span>

|<span data-ttu-id="af631-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-182">Requirement</span></span>| <span data-ttu-id="af631-183">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-185">1,5</span><span class="sxs-lookup"><span data-stu-id="af631-185">1.5</span></span> |
|[<span data-ttu-id="af631-186">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-187">ReadItem</span></span>|
|[<span data-ttu-id="af631-188">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="af631-190">Métodos</span><span class="sxs-lookup"><span data-stu-id="af631-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="af631-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af631-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="af631-192">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="af631-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="af631-193">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="af631-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="af631-194">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="af631-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-195">Parameters</span></span>

| <span data-ttu-id="af631-196">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-196">Name</span></span> | <span data-ttu-id="af631-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-197">Type</span></span> | <span data-ttu-id="af631-198">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-198">Attributes</span></span> | <span data-ttu-id="af631-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="af631-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="af631-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="af631-201">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="af631-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="af631-202">Função</span><span class="sxs-lookup"><span data-stu-id="af631-202">Function</span></span> || <span data-ttu-id="af631-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="af631-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="af631-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-206">Object</span></span> | <span data-ttu-id="af631-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-207">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-208">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af631-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="af631-209">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-209">Object</span></span> | <span data-ttu-id="af631-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-210">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-211">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af631-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="af631-212">function</span><span class="sxs-lookup"><span data-stu-id="af631-212">function</span></span>| <span data-ttu-id="af631-213">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-213">&lt;optional&gt;</span></span>|<span data-ttu-id="af631-214">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-215">Requirements</span></span>

|<span data-ttu-id="af631-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-216">Requirement</span></span>| <span data-ttu-id="af631-217">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-219">1,5</span><span class="sxs-lookup"><span data-stu-id="af631-219">1.5</span></span> |
|[<span data-ttu-id="af631-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-221">ReadItem</span></span> |
|[<span data-ttu-id="af631-222">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-224">Example</span></span>

```javascript
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
};
```

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="af631-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="af631-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="af631-226">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="af631-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-227">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="af631-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-230">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-230">Parameters</span></span>

|<span data-ttu-id="af631-231">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-231">Name</span></span>| <span data-ttu-id="af631-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-232">Type</span></span>| <span data-ttu-id="af631-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="af631-234">String</span><span class="sxs-lookup"><span data-stu-id="af631-234">String</span></span>|<span data-ttu-id="af631-235">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="af631-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="af631-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="af631-237">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="af631-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-238">Requirements</span></span>

|<span data-ttu-id="af631-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-239">Requirement</span></span>| <span data-ttu-id="af631-240">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-242">1.3</span><span class="sxs-lookup"><span data-stu-id="af631-242">1.3</span></span>|
|[<span data-ttu-id="af631-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-244">Restrito</span><span class="sxs-lookup"><span data-stu-id="af631-244">Restricted</span></span>|
|[<span data-ttu-id="af631-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af631-247">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af631-247">Returns:</span></span>

<span data-ttu-id="af631-248">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="af631-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="af631-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="af631-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="af631-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="af631-251">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="af631-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="af631-p108">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="af631-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="af631-p109">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="af631-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-257">Parameters</span></span>

|<span data-ttu-id="af631-258">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-258">Name</span></span>| <span data-ttu-id="af631-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-259">Type</span></span>| <span data-ttu-id="af631-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="af631-261">Date</span><span class="sxs-lookup"><span data-stu-id="af631-261">Date</span></span>|<span data-ttu-id="af631-262">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="af631-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-263">Requirements</span></span>

|<span data-ttu-id="af631-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-264">Requirement</span></span>| <span data-ttu-id="af631-265">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-267">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-267">1.0</span></span>|
|[<span data-ttu-id="af631-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-269">ReadItem</span></span>|
|[<span data-ttu-id="af631-270">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-271">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af631-272">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af631-272">Returns:</span></span>

<span data-ttu-id="af631-273">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="af631-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="af631-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="af631-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="af631-275">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="af631-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-276">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="af631-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-279">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-279">Parameters</span></span>

|<span data-ttu-id="af631-280">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-280">Name</span></span>| <span data-ttu-id="af631-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-281">Type</span></span>| <span data-ttu-id="af631-282">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="af631-283">String</span><span class="sxs-lookup"><span data-stu-id="af631-283">String</span></span>|<span data-ttu-id="af631-284">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="af631-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="af631-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="af631-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="af631-286">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="af631-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-287">Requirements</span></span>

|<span data-ttu-id="af631-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-288">Requirement</span></span>| <span data-ttu-id="af631-289">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-291">1.3</span><span class="sxs-lookup"><span data-stu-id="af631-291">1.3</span></span>|
|[<span data-ttu-id="af631-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-293">Restrito</span><span class="sxs-lookup"><span data-stu-id="af631-293">Restricted</span></span>|
|[<span data-ttu-id="af631-294">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af631-296">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af631-296">Returns:</span></span>

<span data-ttu-id="af631-297">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="af631-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="af631-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="af631-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="af631-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="af631-300">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="af631-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="af631-301">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="af631-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-302">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-302">Parameters</span></span>

|<span data-ttu-id="af631-303">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-303">Name</span></span>| <span data-ttu-id="af631-304">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-304">Type</span></span>| <span data-ttu-id="af631-305">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="af631-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="af631-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="af631-307">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="af631-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-308">Requirements</span></span>

|<span data-ttu-id="af631-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-309">Requirement</span></span>| <span data-ttu-id="af631-310">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-312">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-312">1.0</span></span>|
|[<span data-ttu-id="af631-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-314">ReadItem</span></span>|
|[<span data-ttu-id="af631-315">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="af631-317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="af631-317">Returns:</span></span>

<span data-ttu-id="af631-318">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="af631-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="af631-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="af631-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="af631-320">Date</span><span class="sxs-lookup"><span data-stu-id="af631-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="af631-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="af631-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="af631-322">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="af631-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-323">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-324">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="af631-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="af631-p111">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="af631-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="af631-327">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="af631-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="af631-328">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="af631-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-329">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-329">Parameters</span></span>

|<span data-ttu-id="af631-330">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-330">Name</span></span>| <span data-ttu-id="af631-331">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-331">Type</span></span>| <span data-ttu-id="af631-332">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="af631-333">String</span><span class="sxs-lookup"><span data-stu-id="af631-333">String</span></span>|<span data-ttu-id="af631-334">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="af631-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-335">Requirements</span></span>

|<span data-ttu-id="af631-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-336">Requirement</span></span>| <span data-ttu-id="af631-337">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-339">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-339">1.0</span></span>|
|[<span data-ttu-id="af631-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-341">ReadItem</span></span>|
|[<span data-ttu-id="af631-342">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-343">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="af631-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="af631-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="af631-346">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="af631-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-347">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-348">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="af631-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="af631-349">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af631-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="af631-350">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="af631-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="af631-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="af631-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-353">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-353">Parameters</span></span>

|<span data-ttu-id="af631-354">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-354">Name</span></span>| <span data-ttu-id="af631-355">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-355">Type</span></span>| <span data-ttu-id="af631-356">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="af631-357">String</span><span class="sxs-lookup"><span data-stu-id="af631-357">String</span></span>|<span data-ttu-id="af631-358">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="af631-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-359">Requirements</span></span>

|<span data-ttu-id="af631-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-360">Requirement</span></span>| <span data-ttu-id="af631-361">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-363">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-363">1.0</span></span>|
|[<span data-ttu-id="af631-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-365">ReadItem</span></span>|
|[<span data-ttu-id="af631-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-367">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="af631-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="af631-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="af631-370">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="af631-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-371">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="af631-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="af631-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="af631-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="af631-p114">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="af631-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="af631-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="af631-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="af631-379">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="af631-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-380">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-380">Parameters</span></span>

|<span data-ttu-id="af631-381">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-381">Name</span></span>| <span data-ttu-id="af631-382">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-382">Type</span></span>| <span data-ttu-id="af631-383">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="af631-384">Object</span><span class="sxs-lookup"><span data-stu-id="af631-384">Object</span></span> | <span data-ttu-id="af631-385">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="af631-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="af631-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="af631-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="af631-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="af631-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="af631-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="af631-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="af631-392">Data</span><span class="sxs-lookup"><span data-stu-id="af631-392">Date</span></span> | <span data-ttu-id="af631-393">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="af631-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="af631-394">Data</span><span class="sxs-lookup"><span data-stu-id="af631-394">Date</span></span> | <span data-ttu-id="af631-395">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="af631-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="af631-396">String</span><span class="sxs-lookup"><span data-stu-id="af631-396">String</span></span> | <span data-ttu-id="af631-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af631-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="af631-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="af631-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="af631-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="af631-402">String</span><span class="sxs-lookup"><span data-stu-id="af631-402">String</span></span> | <span data-ttu-id="af631-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="af631-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="af631-405">String</span><span class="sxs-lookup"><span data-stu-id="af631-405">String</span></span> | <span data-ttu-id="af631-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="af631-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="af631-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-408">Requirements</span></span>

|<span data-ttu-id="af631-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-409">Requirement</span></span>| <span data-ttu-id="af631-410">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-412">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-412">1.0</span></span>|
|[<span data-ttu-id="af631-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-414">ReadItem</span></span>|
|[<span data-ttu-id="af631-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-416">Read</span><span class="sxs-lookup"><span data-stu-id="af631-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-417">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="af631-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="af631-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="af631-419">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="af631-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="af631-p122">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="af631-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-422">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="af631-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="af631-423">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="af631-423">**REST Tokens**</span></span>

<span data-ttu-id="af631-p123">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="af631-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="af631-427">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="af631-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="af631-428">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="af631-428">**EWS Tokens**</span></span>

<span data-ttu-id="af631-p124">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="af631-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="af631-431">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="af631-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-432">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-432">Parameters</span></span>

|<span data-ttu-id="af631-433">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-433">Name</span></span>| <span data-ttu-id="af631-434">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-434">Type</span></span>| <span data-ttu-id="af631-435">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-435">Attributes</span></span>| <span data-ttu-id="af631-436">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="af631-437">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-437">Object</span></span> | <span data-ttu-id="af631-438">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-438">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-439">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af631-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="af631-440">Booliano</span><span class="sxs-lookup"><span data-stu-id="af631-440">Boolean</span></span> |  <span data-ttu-id="af631-441">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-441">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-p125">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="af631-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="af631-444">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-444">Object</span></span> |  <span data-ttu-id="af631-445">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-445">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-446">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="af631-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="af631-447">function</span><span class="sxs-lookup"><span data-stu-id="af631-447">function</span></span>||<span data-ttu-id="af631-448">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af631-449">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af631-449">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="af631-450">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="af631-450">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af631-451">Erros</span><span class="sxs-lookup"><span data-stu-id="af631-451">Errors</span></span>

|<span data-ttu-id="af631-452">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af631-452">Error code</span></span>|<span data-ttu-id="af631-453">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-453">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="af631-454">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="af631-454">The request has failed.</span></span> <span data-ttu-id="af631-455">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="af631-455">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="af631-456">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="af631-456">The RMS server returned an error.</span></span> <span data-ttu-id="af631-457">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="af631-457">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="af631-458">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="af631-458">The user is no longer connected to the network.</span></span> <span data-ttu-id="af631-459">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="af631-459">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-460">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-460">Requirements</span></span>

|<span data-ttu-id="af631-461">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-461">Requirement</span></span>| <span data-ttu-id="af631-462">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-463">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-464">1,5</span><span class="sxs-lookup"><span data-stu-id="af631-464">1.5</span></span> |
|[<span data-ttu-id="af631-465">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-466">ReadItem</span></span>|
|[<span data-ttu-id="af631-467">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-468">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="af631-468">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-469">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-469">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="af631-470">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="af631-470">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="af631-471">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="af631-471">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="af631-p129">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="af631-p129">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="af631-p130">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="af631-p130">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="af631-477">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="af631-477">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="af631-p131">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="af631-p131">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-480">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-480">Parameters</span></span>

|<span data-ttu-id="af631-481">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-481">Name</span></span>| <span data-ttu-id="af631-482">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-482">Type</span></span>| <span data-ttu-id="af631-483">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-483">Attributes</span></span>| <span data-ttu-id="af631-484">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-484">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="af631-485">function</span><span class="sxs-lookup"><span data-stu-id="af631-485">function</span></span>||<span data-ttu-id="af631-486">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-486">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af631-487">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af631-487">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="af631-488">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="af631-488">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="af631-489">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-489">Object</span></span>| <span data-ttu-id="af631-490">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-490">&lt;optional&gt;</span></span>|<span data-ttu-id="af631-491">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="af631-491">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af631-492">Erros</span><span class="sxs-lookup"><span data-stu-id="af631-492">Errors</span></span>

|<span data-ttu-id="af631-493">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af631-493">Error code</span></span>|<span data-ttu-id="af631-494">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-494">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="af631-495">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="af631-495">The request has failed.</span></span> <span data-ttu-id="af631-496">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="af631-496">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="af631-497">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="af631-497">The RMS server returned an error.</span></span> <span data-ttu-id="af631-498">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="af631-498">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="af631-499">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="af631-499">The user is no longer connected to the network.</span></span> <span data-ttu-id="af631-500">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="af631-500">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-501">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-501">Requirements</span></span>

|<span data-ttu-id="af631-502">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-502">Requirement</span></span>| <span data-ttu-id="af631-503">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-504">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-505">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-505">1.0</span></span>|
|[<span data-ttu-id="af631-506">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-506">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-507">ReadItem</span></span>|
|[<span data-ttu-id="af631-508">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-508">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-509">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="af631-509">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-510">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-510">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="af631-511">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="af631-511">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="af631-512">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="af631-512">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="af631-513">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="af631-513">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-514">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-514">Parameters</span></span>

|<span data-ttu-id="af631-515">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-515">Name</span></span>| <span data-ttu-id="af631-516">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-516">Type</span></span>| <span data-ttu-id="af631-517">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-517">Attributes</span></span>| <span data-ttu-id="af631-518">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-518">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="af631-519">function</span><span class="sxs-lookup"><span data-stu-id="af631-519">function</span></span>||<span data-ttu-id="af631-520">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af631-521">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af631-521">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="af631-522">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="af631-522">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="af631-523">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-523">Object</span></span>| <span data-ttu-id="af631-524">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-524">&lt;optional&gt;</span></span>|<span data-ttu-id="af631-525">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="af631-525">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="af631-526">Erros</span><span class="sxs-lookup"><span data-stu-id="af631-526">Errors</span></span>

|<span data-ttu-id="af631-527">Código de erro</span><span class="sxs-lookup"><span data-stu-id="af631-527">Error code</span></span>|<span data-ttu-id="af631-528">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-528">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="af631-529">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="af631-529">The request has failed.</span></span> <span data-ttu-id="af631-530">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="af631-530">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="af631-531">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="af631-531">The RMS server returned an error.</span></span> <span data-ttu-id="af631-532">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="af631-532">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="af631-533">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="af631-533">The user is no longer connected to the network.</span></span> <span data-ttu-id="af631-534">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="af631-534">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-535">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-535">Requirements</span></span>

|<span data-ttu-id="af631-536">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-536">Requirement</span></span>| <span data-ttu-id="af631-537">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-538">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-539">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-539">1.0</span></span>|
|[<span data-ttu-id="af631-540">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-541">ReadItem</span></span>|
|[<span data-ttu-id="af631-542">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-543">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-543">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-544">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-544">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="af631-545">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="af631-545">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="af631-546">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="af631-546">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="af631-547">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="af631-547">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="af631-548">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="af631-548">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="af631-549">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="af631-549">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="af631-550">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="af631-550">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="af631-551">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="af631-551">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="af631-552">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="af631-552">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="af631-553">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="af631-553">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="af631-554">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="af631-554">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="af631-p139">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="af631-p139">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="af631-557">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="af631-557">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="af631-558">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="af631-558">Version differences</span></span>

<span data-ttu-id="af631-559">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="af631-559">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="af631-p140">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="af631-p140">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-563">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-563">Parameters</span></span>

|<span data-ttu-id="af631-564">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-564">Name</span></span>| <span data-ttu-id="af631-565">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-565">Type</span></span>| <span data-ttu-id="af631-566">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-566">Attributes</span></span>| <span data-ttu-id="af631-567">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-567">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="af631-568">String</span><span class="sxs-lookup"><span data-stu-id="af631-568">String</span></span>||<span data-ttu-id="af631-569">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="af631-569">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="af631-570">function</span><span class="sxs-lookup"><span data-stu-id="af631-570">function</span></span>||<span data-ttu-id="af631-571">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-571">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="af631-572">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="af631-572">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="af631-573">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="af631-573">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="af631-574">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-574">Object</span></span>| <span data-ttu-id="af631-575">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-575">&lt;optional&gt;</span></span>|<span data-ttu-id="af631-576">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="af631-576">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-577">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-577">Requirements</span></span>

|<span data-ttu-id="af631-578">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-578">Requirement</span></span>| <span data-ttu-id="af631-579">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-580">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-581">1.0</span><span class="sxs-lookup"><span data-stu-id="af631-581">1.0</span></span>|
|[<span data-ttu-id="af631-582">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-583">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="af631-583">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="af631-584">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="af631-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-585">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-585">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af631-586">Exemplo</span><span class="sxs-lookup"><span data-stu-id="af631-586">Example</span></span>

<span data-ttu-id="af631-587">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="af631-587">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="af631-588">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="af631-588">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="af631-589">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="af631-589">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="af631-590">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="af631-590">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="af631-591">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="af631-591">Parameters</span></span>

| <span data-ttu-id="af631-592">Nome</span><span class="sxs-lookup"><span data-stu-id="af631-592">Name</span></span> | <span data-ttu-id="af631-593">Tipo</span><span class="sxs-lookup"><span data-stu-id="af631-593">Type</span></span> | <span data-ttu-id="af631-594">Atributos</span><span class="sxs-lookup"><span data-stu-id="af631-594">Attributes</span></span> | <span data-ttu-id="af631-595">Descrição</span><span class="sxs-lookup"><span data-stu-id="af631-595">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="af631-596">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="af631-596">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="af631-597">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="af631-597">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="af631-598">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-598">Object</span></span> | <span data-ttu-id="af631-599">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-599">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-600">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="af631-600">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="af631-601">Objeto</span><span class="sxs-lookup"><span data-stu-id="af631-601">Object</span></span> | <span data-ttu-id="af631-602">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-602">&lt;optional&gt;</span></span> | <span data-ttu-id="af631-603">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="af631-603">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="af631-604">function</span><span class="sxs-lookup"><span data-stu-id="af631-604">function</span></span>| <span data-ttu-id="af631-605">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="af631-605">&lt;optional&gt;</span></span>|<span data-ttu-id="af631-606">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="af631-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af631-607">Requisitos</span><span class="sxs-lookup"><span data-stu-id="af631-607">Requirements</span></span>

|<span data-ttu-id="af631-608">Requisito</span><span class="sxs-lookup"><span data-stu-id="af631-608">Requirement</span></span>| <span data-ttu-id="af631-609">Valor</span><span class="sxs-lookup"><span data-stu-id="af631-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="af631-610">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="af631-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="af631-611">1,5</span><span class="sxs-lookup"><span data-stu-id="af631-611">1.5</span></span> |
|[<span data-ttu-id="af631-612">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="af631-612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="af631-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="af631-613">ReadItem</span></span> |
|[<span data-ttu-id="af631-614">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="af631-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="af631-615">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="af631-615">Compose or Read</span></span>|
