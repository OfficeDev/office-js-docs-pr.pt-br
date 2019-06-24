---
title: 'Office.context.mailbox: conjunto de requisitos da versão 1.5'
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: a0c1a45fd3eaa9cf324a6854120d642eb7520132
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127279"
---
# <a name="mailbox"></a><span data-ttu-id="4b9fb-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="4b9fb-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4b9fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4b9fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4b9fb-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b9fb-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-105">Requirements</span></span>

|<span data-ttu-id="4b9fb-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-106">Requirement</span></span>| <span data-ttu-id="4b9fb-107">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-109">1.0</span></span>|
|[<span data-ttu-id="4b9fb-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-111">Restricted</span></span>|
|[<span data-ttu-id="4b9fb-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4b9fb-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-114">Members and methods</span></span>

| <span data-ttu-id="4b9fb-115">Membro</span><span class="sxs-lookup"><span data-stu-id="4b9fb-115">Member</span></span> | <span data-ttu-id="4b9fb-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4b9fb-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4b9fb-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4b9fb-118">Membro</span><span class="sxs-lookup"><span data-stu-id="4b9fb-118">Member</span></span> |
| [<span data-ttu-id="4b9fb-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="4b9fb-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="4b9fb-120">Membro</span><span class="sxs-lookup"><span data-stu-id="4b9fb-120">Member</span></span> |
| [<span data-ttu-id="4b9fb-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4b9fb-122">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-122">Method</span></span> |
| [<span data-ttu-id="4b9fb-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="4b9fb-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="4b9fb-124">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-124">Method</span></span> |
| [<span data-ttu-id="4b9fb-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4b9fb-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4b9fb-126">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-126">Method</span></span> |
| [<span data-ttu-id="4b9fb-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="4b9fb-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="4b9fb-128">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-128">Method</span></span> |
| [<span data-ttu-id="4b9fb-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4b9fb-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4b9fb-130">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-130">Method</span></span> |
| [<span data-ttu-id="4b9fb-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4b9fb-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4b9fb-132">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-132">Method</span></span> |
| [<span data-ttu-id="4b9fb-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4b9fb-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4b9fb-134">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-134">Method</span></span> |
| [<span data-ttu-id="4b9fb-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4b9fb-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4b9fb-136">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-136">Method</span></span> |
| [<span data-ttu-id="4b9fb-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="4b9fb-138">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-138">Method</span></span> |
| [<span data-ttu-id="4b9fb-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4b9fb-140">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-140">Method</span></span> |
| [<span data-ttu-id="4b9fb-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4b9fb-142">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-142">Method</span></span> |
| [<span data-ttu-id="4b9fb-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4b9fb-144">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-144">Method</span></span> |
| [<span data-ttu-id="4b9fb-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4b9fb-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4b9fb-146">Método</span><span class="sxs-lookup"><span data-stu-id="4b9fb-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4b9fb-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="4b9fb-147">Namespaces</span></span>

<span data-ttu-id="4b9fb-148">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4b9fb-149">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4b9fb-150">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4b9fb-151">Members</span><span class="sxs-lookup"><span data-stu-id="4b9fb-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4b9fb-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-152">ewsUrl :String</span></span>

<span data-ttu-id="4b9fb-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-155">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4b9fb-158">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4b9fb-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4b9fb-161">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-161">Type</span></span>

*   <span data-ttu-id="4b9fb-162">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b9fb-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-163">Requirements</span></span>

|<span data-ttu-id="4b9fb-164">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-164">Requirement</span></span>| <span data-ttu-id="4b9fb-165">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-166">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-167">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-167">1.0</span></span>|
|[<span data-ttu-id="4b9fb-168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-169">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-170">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="4b9fb-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-172">restUrl :String</span></span>

<span data-ttu-id="4b9fb-173">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="4b9fb-174">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="4b9fb-175">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="4b9fb-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-178">Os clientes do Outlook conectados a instalações locais do Exchange 2016 ou posterior com um REST personalizado da URL configurada retornarão um valor inválido para `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="4b9fb-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-179">Type</span></span>

*   <span data-ttu-id="4b9fb-180">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b9fb-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-181">Requirements</span></span>

|<span data-ttu-id="4b9fb-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-182">Requirement</span></span>| <span data-ttu-id="4b9fb-183">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-185">1,5</span><span class="sxs-lookup"><span data-stu-id="4b9fb-185">1.5</span></span> |
|[<span data-ttu-id="4b9fb-186">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-187">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-188">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4b9fb-190">Métodos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4b9fb-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4b9fb-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4b9fb-192">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4b9fb-193">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="4b9fb-194">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-195">Parameters</span></span>

| <span data-ttu-id="4b9fb-196">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-196">Name</span></span> | <span data-ttu-id="4b9fb-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-197">Type</span></span> | <span data-ttu-id="4b9fb-198">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-198">Attributes</span></span> | <span data-ttu-id="4b9fb-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4b9fb-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4b9fb-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4b9fb-201">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4b9fb-202">Função</span><span class="sxs-lookup"><span data-stu-id="4b9fb-202">Function</span></span> || <span data-ttu-id="4b9fb-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4b9fb-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-206">Object</span></span> | <span data-ttu-id="4b9fb-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-207">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-208">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4b9fb-209">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-209">Object</span></span> | <span data-ttu-id="4b9fb-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-210">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-211">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4b9fb-212">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-212">function</span></span>| <span data-ttu-id="4b9fb-213">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-213">&lt;optional&gt;</span></span>|<span data-ttu-id="4b9fb-214">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-215">Requirements</span></span>

|<span data-ttu-id="4b9fb-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-216">Requirement</span></span>| <span data-ttu-id="4b9fb-217">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-219">1,5</span><span class="sxs-lookup"><span data-stu-id="4b9fb-219">1.5</span></span> |
|[<span data-ttu-id="4b9fb-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-221">ReadItem</span></span> |
|[<span data-ttu-id="4b9fb-222">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4b9fb-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4b9fb-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4b9fb-226">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-227">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-230">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-230">Parameters</span></span>

|<span data-ttu-id="4b9fb-231">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-231">Name</span></span>| <span data-ttu-id="4b9fb-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-232">Type</span></span>| <span data-ttu-id="4b9fb-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4b9fb-234">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-234">String</span></span>|<span data-ttu-id="4b9fb-235">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4b9fb-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4b9fb-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="4b9fb-237">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-238">Requirements</span></span>

|<span data-ttu-id="4b9fb-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-239">Requirement</span></span>| <span data-ttu-id="4b9fb-240">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-242">1.3</span><span class="sxs-lookup"><span data-stu-id="4b9fb-242">1.3</span></span>|
|[<span data-ttu-id="4b9fb-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-244">Restrito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-244">Restricted</span></span>|
|[<span data-ttu-id="4b9fb-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b9fb-247">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4b9fb-247">Returns:</span></span>

<span data-ttu-id="4b9fb-248">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4b9fb-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="4b9fb-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="4b9fb-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="4b9fb-251">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4b9fb-p108">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4b9fb-p109">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-257">Parameters</span></span>

|<span data-ttu-id="4b9fb-258">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-258">Name</span></span>| <span data-ttu-id="4b9fb-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-259">Type</span></span>| <span data-ttu-id="4b9fb-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4b9fb-261">Date</span><span class="sxs-lookup"><span data-stu-id="4b9fb-261">Date</span></span>|<span data-ttu-id="4b9fb-262">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="4b9fb-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-263">Requirements</span></span>

|<span data-ttu-id="4b9fb-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-264">Requirement</span></span>| <span data-ttu-id="4b9fb-265">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-267">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-267">1.0</span></span>|
|[<span data-ttu-id="4b9fb-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-269">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-270">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-271">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b9fb-272">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4b9fb-272">Returns:</span></span>

<span data-ttu-id="4b9fb-273">Tipo: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4b9fb-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4b9fb-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4b9fb-275">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-276">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-279">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-279">Parameters</span></span>

|<span data-ttu-id="4b9fb-280">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-280">Name</span></span>| <span data-ttu-id="4b9fb-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-281">Type</span></span>| <span data-ttu-id="4b9fb-282">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4b9fb-283">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-283">String</span></span>|<span data-ttu-id="4b9fb-284">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4b9fb-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4b9fb-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="4b9fb-286">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-287">Requirements</span></span>

|<span data-ttu-id="4b9fb-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-288">Requirement</span></span>| <span data-ttu-id="4b9fb-289">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-291">1.3</span><span class="sxs-lookup"><span data-stu-id="4b9fb-291">1.3</span></span>|
|[<span data-ttu-id="4b9fb-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-293">Restrito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-293">Restricted</span></span>|
|[<span data-ttu-id="4b9fb-294">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b9fb-296">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4b9fb-296">Returns:</span></span>

<span data-ttu-id="4b9fb-297">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4b9fb-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4b9fb-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4b9fb-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4b9fb-300">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4b9fb-301">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-302">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-302">Parameters</span></span>

|<span data-ttu-id="4b9fb-303">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-303">Name</span></span>| <span data-ttu-id="4b9fb-304">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-304">Type</span></span>| <span data-ttu-id="4b9fb-305">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4b9fb-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4b9fb-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="4b9fb-307">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-308">Requirements</span></span>

|<span data-ttu-id="4b9fb-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-309">Requirement</span></span>| <span data-ttu-id="4b9fb-310">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-312">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-312">1.0</span></span>|
|[<span data-ttu-id="4b9fb-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-314">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-315">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b9fb-317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4b9fb-317">Returns:</span></span>

<span data-ttu-id="4b9fb-318">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="4b9fb-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4b9fb-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4b9fb-320">Date</span><span class="sxs-lookup"><span data-stu-id="4b9fb-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4b9fb-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4b9fb-322">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-323">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-324">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4b9fb-p111">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4b9fb-327">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4b9fb-328">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-329">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-329">Parameters</span></span>

|<span data-ttu-id="4b9fb-330">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-330">Name</span></span>| <span data-ttu-id="4b9fb-331">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-331">Type</span></span>| <span data-ttu-id="4b9fb-332">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4b9fb-333">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-333">String</span></span>|<span data-ttu-id="4b9fb-334">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-335">Requirements</span></span>

|<span data-ttu-id="4b9fb-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-336">Requirement</span></span>| <span data-ttu-id="4b9fb-337">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-339">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-339">1.0</span></span>|
|[<span data-ttu-id="4b9fb-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-341">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-342">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-343">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4b9fb-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4b9fb-346">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-347">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-348">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4b9fb-349">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4b9fb-350">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4b9fb-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-353">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-353">Parameters</span></span>

|<span data-ttu-id="4b9fb-354">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-354">Name</span></span>| <span data-ttu-id="4b9fb-355">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-355">Type</span></span>| <span data-ttu-id="4b9fb-356">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4b9fb-357">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-357">String</span></span>|<span data-ttu-id="4b9fb-358">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-359">Requirements</span></span>

|<span data-ttu-id="4b9fb-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-360">Requirement</span></span>| <span data-ttu-id="4b9fb-361">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-363">1.0</span></span>|
|[<span data-ttu-id="4b9fb-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-365">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-367">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4b9fb-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4b9fb-370">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-371">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4b9fb-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4b9fb-p114">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4b9fb-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4b9fb-379">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-380">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-380">Parameters</span></span>

|<span data-ttu-id="4b9fb-381">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-381">Name</span></span>| <span data-ttu-id="4b9fb-382">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-382">Type</span></span>| <span data-ttu-id="4b9fb-383">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4b9fb-384">Object</span><span class="sxs-lookup"><span data-stu-id="4b9fb-384">Object</span></span> | <span data-ttu-id="4b9fb-385">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4b9fb-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4b9fb-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4b9fb-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4b9fb-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4b9fb-392">Data</span><span class="sxs-lookup"><span data-stu-id="4b9fb-392">Date</span></span> | <span data-ttu-id="4b9fb-393">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4b9fb-394">Data</span><span class="sxs-lookup"><span data-stu-id="4b9fb-394">Date</span></span> | <span data-ttu-id="4b9fb-395">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4b9fb-396">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-396">String</span></span> | <span data-ttu-id="4b9fb-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4b9fb-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4b9fb-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4b9fb-402">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-402">String</span></span> | <span data-ttu-id="4b9fb-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4b9fb-405">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-405">String</span></span> | <span data-ttu-id="4b9fb-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b9fb-408">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-408">Requirements</span></span>

|<span data-ttu-id="4b9fb-409">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-409">Requirement</span></span>| <span data-ttu-id="4b9fb-410">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-411">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-412">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-412">1.0</span></span>|
|[<span data-ttu-id="4b9fb-413">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-414">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-415">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-416">Read</span><span class="sxs-lookup"><span data-stu-id="4b9fb-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-417">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="4b9fb-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4b9fb-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="4b9fb-419">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="4b9fb-p122">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-422">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="4b9fb-423">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="4b9fb-423">**REST Tokens**</span></span>

<span data-ttu-id="4b9fb-p123">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="4b9fb-427">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="4b9fb-428">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="4b9fb-428">**EWS Tokens**</span></span>

<span data-ttu-id="4b9fb-p124">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="4b9fb-431">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-432">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-432">Parameters</span></span>

|<span data-ttu-id="4b9fb-433">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-433">Name</span></span>| <span data-ttu-id="4b9fb-434">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-434">Type</span></span>| <span data-ttu-id="4b9fb-435">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-435">Attributes</span></span>| <span data-ttu-id="4b9fb-436">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="4b9fb-437">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-437">Object</span></span> | <span data-ttu-id="4b9fb-438">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-438">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-439">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="4b9fb-440">Booliano</span><span class="sxs-lookup"><span data-stu-id="4b9fb-440">Boolean</span></span> |  <span data-ttu-id="4b9fb-441">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-441">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-p125">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4b9fb-444">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-444">Object</span></span> |  <span data-ttu-id="4b9fb-445">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-445">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-446">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="4b9fb-447">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-447">function</span></span>||<span data-ttu-id="4b9fb-p126">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-450">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-450">Requirements</span></span>

|<span data-ttu-id="4b9fb-451">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-451">Requirement</span></span>| <span data-ttu-id="4b9fb-452">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-453">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-454">1,5</span><span class="sxs-lookup"><span data-stu-id="4b9fb-454">1.5</span></span> |
|[<span data-ttu-id="4b9fb-455">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-456">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-457">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-458">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="4b9fb-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-459">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4b9fb-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4b9fb-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4b9fb-461">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4b9fb-p127">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p127">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4b9fb-p128">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p128">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4b9fb-467">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="4b9fb-p129">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p129">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-470">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-470">Parameters</span></span>

|<span data-ttu-id="4b9fb-471">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-471">Name</span></span>| <span data-ttu-id="4b9fb-472">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-472">Type</span></span>| <span data-ttu-id="4b9fb-473">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-473">Attributes</span></span>| <span data-ttu-id="4b9fb-474">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4b9fb-475">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-475">function</span></span>||<span data-ttu-id="4b9fb-p130">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p130">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4b9fb-478">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-478">Object</span></span>| <span data-ttu-id="4b9fb-479">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-479">&lt;optional&gt;</span></span>|<span data-ttu-id="4b9fb-480">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-481">Requirements</span></span>

|<span data-ttu-id="4b9fb-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-482">Requirement</span></span>| <span data-ttu-id="4b9fb-483">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-485">1.3</span><span class="sxs-lookup"><span data-stu-id="4b9fb-485">1.3</span></span>|
|[<span data-ttu-id="4b9fb-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-487">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-489">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="4b9fb-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-490">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-490">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4b9fb-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4b9fb-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4b9fb-492">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4b9fb-493">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-494">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-494">Parameters</span></span>

|<span data-ttu-id="4b9fb-495">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-495">Name</span></span>| <span data-ttu-id="4b9fb-496">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-496">Type</span></span>| <span data-ttu-id="4b9fb-497">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-497">Attributes</span></span>| <span data-ttu-id="4b9fb-498">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4b9fb-499">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-499">function</span></span>||<span data-ttu-id="4b9fb-500">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4b9fb-501">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4b9fb-502">Object</span><span class="sxs-lookup"><span data-stu-id="4b9fb-502">Object</span></span>| <span data-ttu-id="4b9fb-503">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-503">&lt;optional&gt;</span></span>|<span data-ttu-id="4b9fb-504">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-505">Requirements</span></span>

|<span data-ttu-id="4b9fb-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-506">Requirement</span></span>| <span data-ttu-id="4b9fb-507">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-509">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-509">1.0</span></span>|
|[<span data-ttu-id="4b9fb-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-511">ReadItem</span></span>|
|[<span data-ttu-id="4b9fb-512">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-513">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-513">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-514">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4b9fb-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4b9fb-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4b9fb-516">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-517">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4b9fb-518">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="4b9fb-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="4b9fb-519">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="4b9fb-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4b9fb-520">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-520">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4b9fb-521">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4b9fb-522">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-522">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4b9fb-523">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4b9fb-524">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4b9fb-p132">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p132">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4b9fb-527">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4b9fb-528">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="4b9fb-528">Version differences</span></span>

<span data-ttu-id="4b9fb-529">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4b9fb-p133">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-p133">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-533">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-533">Parameters</span></span>

|<span data-ttu-id="4b9fb-534">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-534">Name</span></span>| <span data-ttu-id="4b9fb-535">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-535">Type</span></span>| <span data-ttu-id="4b9fb-536">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-536">Attributes</span></span>| <span data-ttu-id="4b9fb-537">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4b9fb-538">String</span><span class="sxs-lookup"><span data-stu-id="4b9fb-538">String</span></span>||<span data-ttu-id="4b9fb-539">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4b9fb-540">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-540">function</span></span>||<span data-ttu-id="4b9fb-541">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4b9fb-542">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4b9fb-543">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4b9fb-544">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-544">Object</span></span>| <span data-ttu-id="4b9fb-545">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-545">&lt;optional&gt;</span></span>|<span data-ttu-id="4b9fb-546">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-547">Requirements</span></span>

|<span data-ttu-id="4b9fb-548">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-548">Requirement</span></span>| <span data-ttu-id="4b9fb-549">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-550">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-551">1.0</span><span class="sxs-lookup"><span data-stu-id="4b9fb-551">1.0</span></span>|
|[<span data-ttu-id="4b9fb-552">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-552">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4b9fb-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4b9fb-554">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4b9fb-554">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-555">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-555">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b9fb-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-556">Example</span></span>

<span data-ttu-id="4b9fb-557">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4b9fb-558">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4b9fb-558">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4b9fb-559">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-559">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4b9fb-560">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b9fb-561">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4b9fb-561">Parameters</span></span>

| <span data-ttu-id="4b9fb-562">Nome</span><span class="sxs-lookup"><span data-stu-id="4b9fb-562">Name</span></span> | <span data-ttu-id="4b9fb-563">Tipo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-563">Type</span></span> | <span data-ttu-id="4b9fb-564">Atributos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-564">Attributes</span></span> | <span data-ttu-id="4b9fb-565">Descrição</span><span class="sxs-lookup"><span data-stu-id="4b9fb-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4b9fb-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4b9fb-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4b9fb-567">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-567">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4b9fb-568">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-568">Object</span></span> | <span data-ttu-id="4b9fb-569">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-569">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-570">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-570">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4b9fb-571">Objeto</span><span class="sxs-lookup"><span data-stu-id="4b9fb-571">Object</span></span> | <span data-ttu-id="4b9fb-572">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-572">&lt;optional&gt;</span></span> | <span data-ttu-id="4b9fb-573">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4b9fb-573">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4b9fb-574">function</span><span class="sxs-lookup"><span data-stu-id="4b9fb-574">function</span></span>| <span data-ttu-id="4b9fb-575">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4b9fb-575">&lt;optional&gt;</span></span>|<span data-ttu-id="4b9fb-576">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b9fb-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b9fb-577">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4b9fb-577">Requirements</span></span>

|<span data-ttu-id="4b9fb-578">Requisito</span><span class="sxs-lookup"><span data-stu-id="4b9fb-578">Requirement</span></span>| <span data-ttu-id="4b9fb-579">Valor</span><span class="sxs-lookup"><span data-stu-id="4b9fb-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b9fb-580">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4b9fb-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b9fb-581">1,5</span><span class="sxs-lookup"><span data-stu-id="4b9fb-581">1.5</span></span> |
|[<span data-ttu-id="4b9fb-582">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4b9fb-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b9fb-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b9fb-583">ReadItem</span></span> |
|[<span data-ttu-id="4b9fb-584">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4b9fb-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b9fb-585">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4b9fb-585">Compose or Read</span></span>|
