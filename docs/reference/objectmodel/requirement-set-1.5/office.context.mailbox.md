---
title: 'Office.context.mailbox: conjunto de requisitos da versão 1.5'
description: ''
ms.date: 11/27/2019
localization_priority: Priority
ms.openlocfilehash: eefeab2cf6fbe78451afae7e588640fe7f50dba4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629683"
---
# <a name="mailbox"></a><span data-ttu-id="03780-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="03780-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="03780-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="03780-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="03780-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="03780-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="03780-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-105">Requirements</span></span>

|<span data-ttu-id="03780-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-106">Requirement</span></span>| <span data-ttu-id="03780-107">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-109">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-109">1.0</span></span>|
|[<span data-ttu-id="03780-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="03780-111">Restricted</span></span>|
|[<span data-ttu-id="03780-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="03780-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="03780-114">Members and methods</span></span>

| <span data-ttu-id="03780-115">Membro</span><span class="sxs-lookup"><span data-stu-id="03780-115">Member</span></span> | <span data-ttu-id="03780-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="03780-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="03780-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="03780-118">Membro</span><span class="sxs-lookup"><span data-stu-id="03780-118">Member</span></span> |
| [<span data-ttu-id="03780-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="03780-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="03780-120">Membro</span><span class="sxs-lookup"><span data-stu-id="03780-120">Member</span></span> |
| [<span data-ttu-id="03780-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="03780-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="03780-122">Método</span><span class="sxs-lookup"><span data-stu-id="03780-122">Method</span></span> |
| [<span data-ttu-id="03780-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="03780-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="03780-124">Método</span><span class="sxs-lookup"><span data-stu-id="03780-124">Method</span></span> |
| [<span data-ttu-id="03780-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="03780-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="03780-126">Método</span><span class="sxs-lookup"><span data-stu-id="03780-126">Method</span></span> |
| [<span data-ttu-id="03780-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="03780-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="03780-128">Método</span><span class="sxs-lookup"><span data-stu-id="03780-128">Method</span></span> |
| [<span data-ttu-id="03780-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="03780-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="03780-130">Método</span><span class="sxs-lookup"><span data-stu-id="03780-130">Method</span></span> |
| [<span data-ttu-id="03780-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="03780-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="03780-132">Método</span><span class="sxs-lookup"><span data-stu-id="03780-132">Method</span></span> |
| [<span data-ttu-id="03780-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="03780-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="03780-134">Método</span><span class="sxs-lookup"><span data-stu-id="03780-134">Method</span></span> |
| [<span data-ttu-id="03780-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="03780-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="03780-136">Método</span><span class="sxs-lookup"><span data-stu-id="03780-136">Method</span></span> |
| [<span data-ttu-id="03780-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03780-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="03780-138">Método</span><span class="sxs-lookup"><span data-stu-id="03780-138">Method</span></span> |
| [<span data-ttu-id="03780-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03780-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="03780-140">Método</span><span class="sxs-lookup"><span data-stu-id="03780-140">Method</span></span> |
| [<span data-ttu-id="03780-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="03780-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="03780-142">Método</span><span class="sxs-lookup"><span data-stu-id="03780-142">Method</span></span> |
| [<span data-ttu-id="03780-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="03780-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="03780-144">Método</span><span class="sxs-lookup"><span data-stu-id="03780-144">Method</span></span> |
| [<span data-ttu-id="03780-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="03780-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="03780-146">Método</span><span class="sxs-lookup"><span data-stu-id="03780-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="03780-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="03780-147">Namespaces</span></span>

<span data-ttu-id="03780-148">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03780-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="03780-149">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03780-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="03780-150">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="03780-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="03780-151">Members</span><span class="sxs-lookup"><span data-stu-id="03780-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="03780-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="03780-152">ewsUrl: String</span></span>

<span data-ttu-id="03780-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="03780-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-155">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03780-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="03780-158">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="03780-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="03780-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="03780-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="03780-161">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-161">Type</span></span>

*   <span data-ttu-id="03780-162">String</span><span class="sxs-lookup"><span data-stu-id="03780-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03780-163">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-163">Requirements</span></span>

|<span data-ttu-id="03780-164">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-164">Requirement</span></span>| <span data-ttu-id="03780-165">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-166">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-167">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-167">1.0</span></span>|
|[<span data-ttu-id="03780-168">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-169">ReadItem</span></span>|
|[<span data-ttu-id="03780-170">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="03780-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="03780-172">restUrl: String</span></span>

<span data-ttu-id="03780-173">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="03780-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="03780-174">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03780-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-175">Os clientes do Outlook conectados a instalações locais do Exchange 2016 ou posterior com um REST personalizado da URL configurada retornarão um valor inválido para `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="03780-175">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="03780-176">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-176">Type</span></span>

*   <span data-ttu-id="03780-177">String</span><span class="sxs-lookup"><span data-stu-id="03780-177">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="03780-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-178">Requirements</span></span>

|<span data-ttu-id="03780-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-179">Requirement</span></span>| <span data-ttu-id="03780-180">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-182">1,5</span><span class="sxs-lookup"><span data-stu-id="03780-182">1.5</span></span> |
|[<span data-ttu-id="03780-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-184">ReadItem</span></span>|
|[<span data-ttu-id="03780-185">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-186">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-186">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="03780-187">Métodos</span><span class="sxs-lookup"><span data-stu-id="03780-187">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="03780-188">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03780-188">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="03780-189">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="03780-189">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="03780-190">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="03780-190">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="03780-191">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="03780-191">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-192">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-192">Parameters</span></span>

| <span data-ttu-id="03780-193">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-193">Name</span></span> | <span data-ttu-id="03780-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-194">Type</span></span> | <span data-ttu-id="03780-195">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-195">Attributes</span></span> | <span data-ttu-id="03780-196">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="03780-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="03780-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="03780-198">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="03780-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="03780-199">Função</span><span class="sxs-lookup"><span data-stu-id="03780-199">Function</span></span> || <span data-ttu-id="03780-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="03780-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="03780-203">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-203">Object</span></span> | <span data-ttu-id="03780-204">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-204">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-205">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03780-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03780-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-206">Object</span></span> | <span data-ttu-id="03780-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-207">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-208">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="03780-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="03780-209">function</span><span class="sxs-lookup"><span data-stu-id="03780-209">function</span></span>| <span data-ttu-id="03780-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-210">&lt;optional&gt;</span></span>|<span data-ttu-id="03780-211">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-212">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-212">Requirements</span></span>

|<span data-ttu-id="03780-213">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-213">Requirement</span></span>| <span data-ttu-id="03780-214">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-215">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-216">1,5</span><span class="sxs-lookup"><span data-stu-id="03780-216">1.5</span></span> |
|[<span data-ttu-id="03780-217">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-218">ReadItem</span></span> |
|[<span data-ttu-id="03780-219">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-221">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-221">Example</span></span>

```js
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

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="03780-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="03780-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="03780-223">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="03780-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-224">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="03780-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-227">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-227">Parameters</span></span>

|<span data-ttu-id="03780-228">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-228">Name</span></span>| <span data-ttu-id="03780-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-229">Type</span></span>| <span data-ttu-id="03780-230">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03780-231">String</span><span class="sxs-lookup"><span data-stu-id="03780-231">String</span></span>|<span data-ttu-id="03780-232">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="03780-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="03780-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="03780-234">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="03780-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-235">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-235">Requirements</span></span>

|<span data-ttu-id="03780-236">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-236">Requirement</span></span>| <span data-ttu-id="03780-237">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-238">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-239">1.3</span><span class="sxs-lookup"><span data-stu-id="03780-239">1.3</span></span>|
|[<span data-ttu-id="03780-240">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-241">Restrito</span><span class="sxs-lookup"><span data-stu-id="03780-241">Restricted</span></span>|
|[<span data-ttu-id="03780-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03780-244">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03780-244">Returns:</span></span>

<span data-ttu-id="03780-245">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="03780-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="03780-246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="03780-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="03780-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="03780-248">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="03780-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="03780-p107">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="03780-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="03780-p108">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="03780-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-254">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-254">Parameters</span></span>

|<span data-ttu-id="03780-255">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-255">Name</span></span>| <span data-ttu-id="03780-256">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-256">Type</span></span>| <span data-ttu-id="03780-257">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="03780-258">Date</span><span class="sxs-lookup"><span data-stu-id="03780-258">Date</span></span>|<span data-ttu-id="03780-259">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="03780-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-260">Requirements</span></span>

|<span data-ttu-id="03780-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-261">Requirement</span></span>| <span data-ttu-id="03780-262">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-264">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-264">1.0</span></span>|
|[<span data-ttu-id="03780-265">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-266">ReadItem</span></span>|
|[<span data-ttu-id="03780-267">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-268">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03780-269">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03780-269">Returns:</span></span>

<span data-ttu-id="03780-270">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="03780-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="03780-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="03780-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="03780-272">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="03780-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-273">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="03780-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-276">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-276">Parameters</span></span>

|<span data-ttu-id="03780-277">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-277">Name</span></span>| <span data-ttu-id="03780-278">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-278">Type</span></span>| <span data-ttu-id="03780-279">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03780-280">String</span><span class="sxs-lookup"><span data-stu-id="03780-280">String</span></span>|<span data-ttu-id="03780-281">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="03780-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="03780-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="03780-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="03780-283">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="03780-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-284">Requirements</span></span>

|<span data-ttu-id="03780-285">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-285">Requirement</span></span>| <span data-ttu-id="03780-286">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-287">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-288">1.3</span><span class="sxs-lookup"><span data-stu-id="03780-288">1.3</span></span>|
|[<span data-ttu-id="03780-289">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-290">Restrito</span><span class="sxs-lookup"><span data-stu-id="03780-290">Restricted</span></span>|
|[<span data-ttu-id="03780-291">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-292">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03780-293">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03780-293">Returns:</span></span>

<span data-ttu-id="03780-294">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="03780-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="03780-295">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="03780-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="03780-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="03780-297">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="03780-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="03780-298">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="03780-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-299">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-299">Parameters</span></span>

|<span data-ttu-id="03780-300">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-300">Name</span></span>| <span data-ttu-id="03780-301">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-301">Type</span></span>| <span data-ttu-id="03780-302">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="03780-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="03780-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="03780-304">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="03780-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-305">Requirements</span></span>

|<span data-ttu-id="03780-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-306">Requirement</span></span>| <span data-ttu-id="03780-307">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-309">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-309">1.0</span></span>|
|[<span data-ttu-id="03780-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-311">ReadItem</span></span>|
|[<span data-ttu-id="03780-312">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-313">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="03780-314">Retorna:</span><span class="sxs-lookup"><span data-stu-id="03780-314">Returns:</span></span>

<span data-ttu-id="03780-315">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="03780-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="03780-316">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="03780-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="03780-317">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-317">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="03780-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="03780-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="03780-319">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="03780-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-320">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-321">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="03780-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="03780-p110">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="03780-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="03780-324">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="03780-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="03780-325">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="03780-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-326">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-326">Parameters</span></span>

|<span data-ttu-id="03780-327">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-327">Name</span></span>| <span data-ttu-id="03780-328">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-328">Type</span></span>| <span data-ttu-id="03780-329">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03780-330">String</span><span class="sxs-lookup"><span data-stu-id="03780-330">String</span></span>|<span data-ttu-id="03780-331">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="03780-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-332">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-332">Requirements</span></span>

|<span data-ttu-id="03780-333">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-333">Requirement</span></span>| <span data-ttu-id="03780-334">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-335">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-336">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-336">1.0</span></span>|
|[<span data-ttu-id="03780-337">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-338">ReadItem</span></span>|
|[<span data-ttu-id="03780-339">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-340">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-341">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="03780-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="03780-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="03780-343">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="03780-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-344">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-345">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="03780-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="03780-346">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="03780-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="03780-347">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="03780-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="03780-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="03780-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-350">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-350">Parameters</span></span>

|<span data-ttu-id="03780-351">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-351">Name</span></span>| <span data-ttu-id="03780-352">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-352">Type</span></span>| <span data-ttu-id="03780-353">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="03780-354">String</span><span class="sxs-lookup"><span data-stu-id="03780-354">String</span></span>|<span data-ttu-id="03780-355">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="03780-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-356">Requirements</span></span>

|<span data-ttu-id="03780-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-357">Requirement</span></span>| <span data-ttu-id="03780-358">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-359">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-360">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-360">1.0</span></span>|
|[<span data-ttu-id="03780-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-362">ReadItem</span></span>|
|[<span data-ttu-id="03780-363">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-364">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-365">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="03780-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="03780-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="03780-367">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="03780-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-368">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="03780-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="03780-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="03780-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="03780-p113">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="03780-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="03780-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="03780-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="03780-376">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="03780-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-377">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-377">Parameters</span></span>

|<span data-ttu-id="03780-378">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-378">Name</span></span>| <span data-ttu-id="03780-379">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-379">Type</span></span>| <span data-ttu-id="03780-380">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-380">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="03780-381">Object</span><span class="sxs-lookup"><span data-stu-id="03780-381">Object</span></span> | <span data-ttu-id="03780-382">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="03780-382">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="03780-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="03780-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03780-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="03780-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="03780-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03780-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="03780-389">Data</span><span class="sxs-lookup"><span data-stu-id="03780-389">Date</span></span> | <span data-ttu-id="03780-390">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="03780-390">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="03780-391">Data</span><span class="sxs-lookup"><span data-stu-id="03780-391">Date</span></span> | <span data-ttu-id="03780-392">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="03780-392">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="03780-393">String</span><span class="sxs-lookup"><span data-stu-id="03780-393">String</span></span> | <span data-ttu-id="03780-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03780-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="03780-396">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-396">Array.&lt;String&gt;</span></span> | <span data-ttu-id="03780-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="03780-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="03780-399">String</span><span class="sxs-lookup"><span data-stu-id="03780-399">String</span></span> | <span data-ttu-id="03780-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="03780-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="03780-402">String</span><span class="sxs-lookup"><span data-stu-id="03780-402">String</span></span> | <span data-ttu-id="03780-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="03780-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="03780-405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-405">Requirements</span></span>

|<span data-ttu-id="03780-406">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-406">Requirement</span></span>| <span data-ttu-id="03780-407">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-408">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-409">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-409">1.0</span></span>|
|[<span data-ttu-id="03780-410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-411">ReadItem</span></span>|
|[<span data-ttu-id="03780-412">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-413">Read</span><span class="sxs-lookup"><span data-stu-id="03780-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-414">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-414">Example</span></span>

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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="03780-415">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="03780-415">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="03780-416">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="03780-416">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="03780-p121">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="03780-p121">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-419">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="03780-419">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="03780-420">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="03780-420">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="03780-421">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="03780-421">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="03780-422">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="03780-422">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="03780-423">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="03780-423">**REST Tokens**</span></span>

<span data-ttu-id="03780-p123">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="03780-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="03780-427">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="03780-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="03780-428">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="03780-428">**EWS Tokens**</span></span>

<span data-ttu-id="03780-p124">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="03780-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="03780-431">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="03780-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="03780-432">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="03780-432">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="03780-433">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="03780-433">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="03780-434">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03780-434">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-435">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-435">Parameters</span></span>

|<span data-ttu-id="03780-436">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-436">Name</span></span>| <span data-ttu-id="03780-437">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-437">Type</span></span>| <span data-ttu-id="03780-438">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-438">Attributes</span></span>| <span data-ttu-id="03780-439">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-439">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="03780-440">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-440">Object</span></span> | <span data-ttu-id="03780-441">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-441">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-442">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03780-442">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="03780-443">Booliano</span><span class="sxs-lookup"><span data-stu-id="03780-443">Boolean</span></span> |  <span data-ttu-id="03780-444">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-444">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-p126">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="03780-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03780-447">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-447">Object</span></span> |  <span data-ttu-id="03780-448">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-448">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-449">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03780-449">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="03780-450">function</span><span class="sxs-lookup"><span data-stu-id="03780-450">function</span></span>||<span data-ttu-id="03780-451">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-451">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03780-452">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03780-452">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03780-453">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03780-453">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03780-454">Erros</span><span class="sxs-lookup"><span data-stu-id="03780-454">Errors</span></span>

|<span data-ttu-id="03780-455">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03780-455">Error code</span></span>|<span data-ttu-id="03780-456">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-456">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03780-457">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03780-457">The request has failed.</span></span> <span data-ttu-id="03780-458">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03780-458">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03780-459">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03780-459">The Exchange server returned an error.</span></span> <span data-ttu-id="03780-460">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03780-460">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03780-461">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03780-461">The user is no longer connected to the network.</span></span> <span data-ttu-id="03780-462">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03780-462">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-463">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-463">Requirements</span></span>

|<span data-ttu-id="03780-464">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-464">Requirement</span></span>| <span data-ttu-id="03780-465">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-466">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-467">1,5</span><span class="sxs-lookup"><span data-stu-id="03780-467">1.5</span></span> |
|[<span data-ttu-id="03780-468">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-469">ReadItem</span></span>|
|[<span data-ttu-id="03780-470">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-471">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="03780-471">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-472">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-472">Example</span></span>

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

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="03780-473">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03780-473">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="03780-474">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="03780-474">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="03780-p130">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="03780-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="03780-477">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="03780-477">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="03780-478">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="03780-478">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="03780-479">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="03780-479">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="03780-480">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="03780-480">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="03780-481">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="03780-481">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="03780-482">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="03780-482">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-483">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-483">Parameters</span></span>

|<span data-ttu-id="03780-484">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-484">Name</span></span>| <span data-ttu-id="03780-485">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-485">Type</span></span>| <span data-ttu-id="03780-486">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-486">Attributes</span></span>| <span data-ttu-id="03780-487">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-487">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="03780-488">function</span><span class="sxs-lookup"><span data-stu-id="03780-488">function</span></span>||<span data-ttu-id="03780-489">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-489">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03780-490">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03780-490">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03780-491">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03780-491">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="03780-492">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-492">Object</span></span>| <span data-ttu-id="03780-493">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-493">&lt;optional&gt;</span></span>|<span data-ttu-id="03780-494">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03780-494">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03780-495">Erros</span><span class="sxs-lookup"><span data-stu-id="03780-495">Errors</span></span>

|<span data-ttu-id="03780-496">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03780-496">Error code</span></span>|<span data-ttu-id="03780-497">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-497">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03780-498">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03780-498">The request has failed.</span></span> <span data-ttu-id="03780-499">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03780-499">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03780-500">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03780-500">The Exchange server returned an error.</span></span> <span data-ttu-id="03780-501">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03780-501">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03780-502">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03780-502">The user is no longer connected to the network.</span></span> <span data-ttu-id="03780-503">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03780-503">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-504">Requirements</span></span>

|<span data-ttu-id="03780-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-505">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="03780-506">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-507">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-507">1.0</span></span> | <span data-ttu-id="03780-508">1.3</span><span class="sxs-lookup"><span data-stu-id="03780-508">1.3</span></span> |
|[<span data-ttu-id="03780-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-510">ReadItem</span></span> | <span data-ttu-id="03780-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-511">ReadItem</span></span> |
|[<span data-ttu-id="03780-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-513">Read</span><span class="sxs-lookup"><span data-stu-id="03780-513">Read</span></span> | <span data-ttu-id="03780-514">Escrever</span><span class="sxs-lookup"><span data-stu-id="03780-514">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="03780-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-515">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="03780-516">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03780-516">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="03780-517">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="03780-517">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="03780-518">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="03780-518">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-519">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-519">Parameters</span></span>

|<span data-ttu-id="03780-520">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-520">Name</span></span>| <span data-ttu-id="03780-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-521">Type</span></span>| <span data-ttu-id="03780-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-522">Attributes</span></span>| <span data-ttu-id="03780-523">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-523">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="03780-524">function</span><span class="sxs-lookup"><span data-stu-id="03780-524">function</span></span>||<span data-ttu-id="03780-525">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-525">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03780-526">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03780-526">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="03780-527">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="03780-527">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="03780-528">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-528">Object</span></span>| <span data-ttu-id="03780-529">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-529">&lt;optional&gt;</span></span>|<span data-ttu-id="03780-530">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03780-530">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="03780-531">Erros</span><span class="sxs-lookup"><span data-stu-id="03780-531">Errors</span></span>

|<span data-ttu-id="03780-532">Código de erro</span><span class="sxs-lookup"><span data-stu-id="03780-532">Error code</span></span>|<span data-ttu-id="03780-533">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-533">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="03780-534">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="03780-534">The request has failed.</span></span> <span data-ttu-id="03780-535">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="03780-535">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="03780-536">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="03780-536">The Exchange server returned an error.</span></span> <span data-ttu-id="03780-537">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="03780-537">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="03780-538">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="03780-538">The user is no longer connected to the network.</span></span> <span data-ttu-id="03780-539">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="03780-539">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-540">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-540">Requirements</span></span>

|<span data-ttu-id="03780-541">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-541">Requirement</span></span>| <span data-ttu-id="03780-542">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-543">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-544">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-544">1.0</span></span>|
|[<span data-ttu-id="03780-545">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-545">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-546">ReadItem</span></span>|
|[<span data-ttu-id="03780-547">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-547">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-548">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-548">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-549">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-549">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="03780-550">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="03780-550">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="03780-551">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03780-551">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="03780-552">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="03780-552">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="03780-553">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="03780-553">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="03780-554">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="03780-554">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="03780-555">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="03780-555">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="03780-556">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="03780-556">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="03780-557">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="03780-557">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="03780-558">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="03780-558">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="03780-559">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="03780-559">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="03780-p140">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="03780-p140">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="03780-562">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="03780-562">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="03780-563">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="03780-563">Version differences</span></span>

<span data-ttu-id="03780-564">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="03780-564">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="03780-p141">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="03780-p141">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-568">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-568">Parameters</span></span>

|<span data-ttu-id="03780-569">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-569">Name</span></span>| <span data-ttu-id="03780-570">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-570">Type</span></span>| <span data-ttu-id="03780-571">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-571">Attributes</span></span>| <span data-ttu-id="03780-572">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-572">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="03780-573">String</span><span class="sxs-lookup"><span data-stu-id="03780-573">String</span></span>||<span data-ttu-id="03780-574">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="03780-574">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="03780-575">function</span><span class="sxs-lookup"><span data-stu-id="03780-575">function</span></span>||<span data-ttu-id="03780-576">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="03780-577">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="03780-577">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="03780-578">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="03780-578">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="03780-579">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-579">Object</span></span>| <span data-ttu-id="03780-580">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-580">&lt;optional&gt;</span></span>|<span data-ttu-id="03780-581">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="03780-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-582">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-582">Requirements</span></span>

|<span data-ttu-id="03780-583">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-583">Requirement</span></span>| <span data-ttu-id="03780-584">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-585">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-586">1.0</span><span class="sxs-lookup"><span data-stu-id="03780-586">1.0</span></span>|
|[<span data-ttu-id="03780-587">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-588">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="03780-588">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="03780-589">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="03780-589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-590">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-590">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="03780-591">Exemplo</span><span class="sxs-lookup"><span data-stu-id="03780-591">Example</span></span>

<span data-ttu-id="03780-592">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="03780-592">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="03780-593">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="03780-593">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="03780-594">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="03780-594">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="03780-595">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="03780-595">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="03780-596">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="03780-596">Parameters</span></span>

| <span data-ttu-id="03780-597">Nome</span><span class="sxs-lookup"><span data-stu-id="03780-597">Name</span></span> | <span data-ttu-id="03780-598">Tipo</span><span class="sxs-lookup"><span data-stu-id="03780-598">Type</span></span> | <span data-ttu-id="03780-599">Atributos</span><span class="sxs-lookup"><span data-stu-id="03780-599">Attributes</span></span> | <span data-ttu-id="03780-600">Descrição</span><span class="sxs-lookup"><span data-stu-id="03780-600">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="03780-601">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="03780-601">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="03780-602">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="03780-602">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="03780-603">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-603">Object</span></span> | <span data-ttu-id="03780-604">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-604">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-605">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="03780-605">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="03780-606">Objeto</span><span class="sxs-lookup"><span data-stu-id="03780-606">Object</span></span> | <span data-ttu-id="03780-607">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-607">&lt;optional&gt;</span></span> | <span data-ttu-id="03780-608">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="03780-608">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="03780-609">function</span><span class="sxs-lookup"><span data-stu-id="03780-609">function</span></span>| <span data-ttu-id="03780-610">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="03780-610">&lt;optional&gt;</span></span>|<span data-ttu-id="03780-611">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="03780-611">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="03780-612">Requisitos</span><span class="sxs-lookup"><span data-stu-id="03780-612">Requirements</span></span>

|<span data-ttu-id="03780-613">Requisito</span><span class="sxs-lookup"><span data-stu-id="03780-613">Requirement</span></span>| <span data-ttu-id="03780-614">Valor</span><span class="sxs-lookup"><span data-stu-id="03780-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="03780-615">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="03780-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="03780-616">1,5</span><span class="sxs-lookup"><span data-stu-id="03780-616">1.5</span></span> |
|[<span data-ttu-id="03780-617">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="03780-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="03780-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="03780-618">ReadItem</span></span> |
|[<span data-ttu-id="03780-619">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="03780-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="03780-620">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="03780-620">Compose or Read</span></span>|
