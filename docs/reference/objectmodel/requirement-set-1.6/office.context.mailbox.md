---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5e85610c93d1001f0a866afa90689c172387fbbc
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127307"
---
# <a name="mailbox"></a><span data-ttu-id="fabfb-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="fabfb-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="fabfb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="fabfb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="fabfb-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="fabfb-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fabfb-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-105">Requirements</span></span>

|<span data-ttu-id="fabfb-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-106">Requirement</span></span>| <span data-ttu-id="fabfb-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-109">1.0</span></span>|
|[<span data-ttu-id="fabfb-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="fabfb-111">Restricted</span></span>|
|[<span data-ttu-id="fabfb-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fabfb-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="fabfb-114">Members and methods</span></span>

| <span data-ttu-id="fabfb-115">Membro</span><span class="sxs-lookup"><span data-stu-id="fabfb-115">Member</span></span> | <span data-ttu-id="fabfb-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fabfb-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="fabfb-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="fabfb-118">Membro</span><span class="sxs-lookup"><span data-stu-id="fabfb-118">Member</span></span> |
| [<span data-ttu-id="fabfb-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="fabfb-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="fabfb-120">Membro</span><span class="sxs-lookup"><span data-stu-id="fabfb-120">Member</span></span> |
| [<span data-ttu-id="fabfb-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="fabfb-122">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-122">Method</span></span> |
| [<span data-ttu-id="fabfb-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="fabfb-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="fabfb-124">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-124">Method</span></span> |
| [<span data-ttu-id="fabfb-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="fabfb-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="fabfb-126">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-126">Method</span></span> |
| [<span data-ttu-id="fabfb-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="fabfb-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="fabfb-128">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-128">Method</span></span> |
| [<span data-ttu-id="fabfb-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="fabfb-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="fabfb-130">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-130">Method</span></span> |
| [<span data-ttu-id="fabfb-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="fabfb-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="fabfb-132">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-132">Method</span></span> |
| [<span data-ttu-id="fabfb-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="fabfb-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="fabfb-134">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-134">Method</span></span> |
| [<span data-ttu-id="fabfb-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="fabfb-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="fabfb-136">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-136">Method</span></span> |
| [<span data-ttu-id="fabfb-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="fabfb-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="fabfb-138">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-138">Method</span></span> |
| [<span data-ttu-id="fabfb-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="fabfb-140">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-140">Method</span></span> |
| [<span data-ttu-id="fabfb-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="fabfb-142">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-142">Method</span></span> |
| [<span data-ttu-id="fabfb-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="fabfb-144">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-144">Method</span></span> |
| [<span data-ttu-id="fabfb-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="fabfb-146">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-146">Method</span></span> |
| [<span data-ttu-id="fabfb-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fabfb-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="fabfb-148">Método</span><span class="sxs-lookup"><span data-stu-id="fabfb-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fabfb-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="fabfb-149">Namespaces</span></span>

<span data-ttu-id="fabfb-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fabfb-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="fabfb-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fabfb-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="fabfb-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fabfb-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="fabfb-153">Membros</span><span class="sxs-lookup"><span data-stu-id="fabfb-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="fabfb-154">ewsUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fabfb-154">ewsUrl: String</span></span>

<span data-ttu-id="fabfb-155">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="fabfb-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="fabfb-156">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fabfb-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-157">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="fabfb-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="fabfb-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fabfb-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="fabfb-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="fabfb-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-163">Type</span></span>

*   <span data-ttu-id="fabfb-164">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fabfb-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-165">Requirements</span></span>

|<span data-ttu-id="fabfb-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-166">Requirement</span></span>| <span data-ttu-id="fabfb-167">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-169">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-169">1.0</span></span>|
|[<span data-ttu-id="fabfb-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-171">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="fabfb-174">restUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fabfb-174">restUrl: String</span></span>

<span data-ttu-id="fabfb-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="fabfb-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="fabfb-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="fabfb-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="fabfb-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fabfb-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="fabfb-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="fabfb-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-180">Type</span></span>

*   <span data-ttu-id="fabfb-181">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fabfb-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-182">Requirements</span></span>

|<span data-ttu-id="fabfb-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-183">Requirement</span></span>| <span data-ttu-id="fabfb-184">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-186">1,5</span><span class="sxs-lookup"><span data-stu-id="fabfb-186">1.5</span></span> |
|[<span data-ttu-id="fabfb-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-188">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="fabfb-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="fabfb-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="fabfb-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fabfb-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="fabfb-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="fabfb-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="fabfb-194">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="fabfb-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="fabfb-195">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="fabfb-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-196">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-196">Parameters</span></span>

| <span data-ttu-id="fabfb-197">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-197">Name</span></span> | <span data-ttu-id="fabfb-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-198">Type</span></span> | <span data-ttu-id="fabfb-199">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-199">Attributes</span></span> | <span data-ttu-id="fabfb-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fabfb-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fabfb-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fabfb-202">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="fabfb-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="fabfb-203">Função</span><span class="sxs-lookup"><span data-stu-id="fabfb-203">Function</span></span> || <span data-ttu-id="fabfb-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="fabfb-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-207">Object</span></span> | <span data-ttu-id="fabfb-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-208">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-209">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fabfb-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fabfb-210">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-210">Object</span></span> | <span data-ttu-id="fabfb-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-211">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-212">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fabfb-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fabfb-213">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-213">function</span></span>| <span data-ttu-id="fabfb-214">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-214">&lt;optional&gt;</span></span>|<span data-ttu-id="fabfb-215">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fabfb-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-216">Requirements</span></span>

|<span data-ttu-id="fabfb-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-217">Requirement</span></span>| <span data-ttu-id="fabfb-218">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-220">1,5</span><span class="sxs-lookup"><span data-stu-id="fabfb-220">1.5</span></span> |
|[<span data-ttu-id="fabfb-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-222">ReadItem</span></span> |
|[<span data-ttu-id="fabfb-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-225">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="fabfb-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="fabfb-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="fabfb-227">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="fabfb-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-228">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-231">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-231">Parameters</span></span>

|<span data-ttu-id="fabfb-232">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-232">Name</span></span>| <span data-ttu-id="fabfb-233">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-233">Type</span></span>| <span data-ttu-id="fabfb-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fabfb-235">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-235">String</span></span>|<span data-ttu-id="fabfb-236">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="fabfb-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="fabfb-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="fabfb-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="fabfb-238">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="fabfb-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-239">Requirements</span></span>

|<span data-ttu-id="fabfb-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-240">Requirement</span></span>| <span data-ttu-id="fabfb-241">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-243">1.3</span><span class="sxs-lookup"><span data-stu-id="fabfb-243">1.3</span></span>|
|[<span data-ttu-id="fabfb-244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-245">Restrito</span><span class="sxs-lookup"><span data-stu-id="fabfb-245">Restricted</span></span>|
|[<span data-ttu-id="fabfb-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fabfb-248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fabfb-248">Returns:</span></span>

<span data-ttu-id="fabfb-249">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="fabfb-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="fabfb-250">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="fabfb-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="fabfb-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="fabfb-252">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="fabfb-253">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para datas e horas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="fabfb-254">O Outlook em uma área de trabalho usa o fuso horário do computador cliente; O Outlook na Web usa o fuso horário definido no centro de administração do Exchange (Eat).</span><span class="sxs-lookup"><span data-stu-id="fabfb-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="fabfb-255">Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="fabfb-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="fabfb-256">Se o aplicativo de email estiver em execução no Outlook em um cliente desktop `convertToLocalClientTime` , o método retornará um objeto Dictionary com os valores definidos para o fuso horário do computador cliente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="fabfb-257">Se o aplicativo de email estiver em execução no Outlook na Web, `convertToLocalClientTime` o método retornará um objeto Dictionary com os valores definidos para o fuso horário especificado no Eat.</span><span class="sxs-lookup"><span data-stu-id="fabfb-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-258">Parameters</span></span>

|<span data-ttu-id="fabfb-259">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-259">Name</span></span>| <span data-ttu-id="fabfb-260">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-260">Type</span></span>| <span data-ttu-id="fabfb-261">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="fabfb-262">Date</span><span class="sxs-lookup"><span data-stu-id="fabfb-262">Date</span></span>|<span data-ttu-id="fabfb-263">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="fabfb-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-264">Requirements</span></span>

|<span data-ttu-id="fabfb-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-265">Requirement</span></span>| <span data-ttu-id="fabfb-266">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-268">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-268">1.0</span></span>|
|[<span data-ttu-id="fabfb-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-270">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fabfb-273">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fabfb-273">Returns:</span></span>

<span data-ttu-id="fabfb-274">Tipo: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="fabfb-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="fabfb-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="fabfb-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="fabfb-276">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="fabfb-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-277">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-280">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-280">Parameters</span></span>

|<span data-ttu-id="fabfb-281">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-281">Name</span></span>| <span data-ttu-id="fabfb-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-282">Type</span></span>| <span data-ttu-id="fabfb-283">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fabfb-284">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-284">String</span></span>|<span data-ttu-id="fabfb-285">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="fabfb-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="fabfb-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="fabfb-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="fabfb-287">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="fabfb-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-288">Requirements</span></span>

|<span data-ttu-id="fabfb-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-289">Requirement</span></span>| <span data-ttu-id="fabfb-290">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-292">1.3</span><span class="sxs-lookup"><span data-stu-id="fabfb-292">1.3</span></span>|
|[<span data-ttu-id="fabfb-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-294">Restrito</span><span class="sxs-lookup"><span data-stu-id="fabfb-294">Restricted</span></span>|
|[<span data-ttu-id="fabfb-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-296">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fabfb-297">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fabfb-297">Returns:</span></span>

<span data-ttu-id="fabfb-298">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="fabfb-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="fabfb-299">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="fabfb-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="fabfb-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="fabfb-301">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="fabfb-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="fabfb-302">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="fabfb-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-303">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-303">Parameters</span></span>

|<span data-ttu-id="fabfb-304">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-304">Name</span></span>| <span data-ttu-id="fabfb-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-305">Type</span></span>| <span data-ttu-id="fabfb-306">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="fabfb-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="fabfb-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="fabfb-308">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="fabfb-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-309">Requirements</span></span>

|<span data-ttu-id="fabfb-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-310">Requirement</span></span>| <span data-ttu-id="fabfb-311">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-313">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-313">1.0</span></span>|
|[<span data-ttu-id="fabfb-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-315">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-317">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fabfb-318">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fabfb-318">Returns:</span></span>

<span data-ttu-id="fabfb-319">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="fabfb-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="fabfb-320">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fabfb-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fabfb-321">Date</span><span class="sxs-lookup"><span data-stu-id="fabfb-321">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="fabfb-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="fabfb-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="fabfb-323">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-324">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-325">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="fabfb-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="fabfb-326">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso mestre de uma série recorrente, mas não é possível exibir uma instância da série.</span><span class="sxs-lookup"><span data-stu-id="fabfb-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="fabfb-327">Isso ocorre porque, no Outlook no Mac, você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="fabfb-328">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fabfb-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="fabfb-329">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="fabfb-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-330">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-330">Parameters</span></span>

|<span data-ttu-id="fabfb-331">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-331">Name</span></span>| <span data-ttu-id="fabfb-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-332">Type</span></span>| <span data-ttu-id="fabfb-333">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fabfb-334">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-334">String</span></span>|<span data-ttu-id="fabfb-335">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-336">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-336">Requirements</span></span>

|<span data-ttu-id="fabfb-337">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-337">Requirement</span></span>| <span data-ttu-id="fabfb-338">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-339">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-340">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-340">1.0</span></span>|
|[<span data-ttu-id="fabfb-341">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-342">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-343">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-344">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-345">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="fabfb-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="fabfb-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="fabfb-347">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-348">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-349">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="fabfb-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="fabfb-350">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fabfb-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="fabfb-351">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="fabfb-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="fabfb-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-354">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-354">Parameters</span></span>

|<span data-ttu-id="fabfb-355">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-355">Name</span></span>| <span data-ttu-id="fabfb-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-356">Type</span></span>| <span data-ttu-id="fabfb-357">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="fabfb-358">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-358">String</span></span>|<span data-ttu-id="fabfb-359">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="fabfb-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-360">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-360">Requirements</span></span>

|<span data-ttu-id="fabfb-361">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-361">Requirement</span></span>| <span data-ttu-id="fabfb-362">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-364">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-364">1.0</span></span>|
|[<span data-ttu-id="fabfb-365">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-366">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-367">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="fabfb-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-368">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-369">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="fabfb-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="fabfb-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="fabfb-371">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="fabfb-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-372">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="fabfb-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="fabfb-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="fabfb-375">No Outlook na Web e dispositivos móveis, este método sempre exibe um formulário com um campo participantes.</span><span class="sxs-lookup"><span data-stu-id="fabfb-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="fabfb-376">Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="fabfb-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="fabfb-377">Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="fabfb-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="fabfb-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="fabfb-380">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fabfb-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-381">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-382">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="fabfb-382">All parameters are optional.</span></span>

|<span data-ttu-id="fabfb-383">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-383">Name</span></span>| <span data-ttu-id="fabfb-384">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-384">Type</span></span>| <span data-ttu-id="fabfb-385">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="fabfb-386">Object</span><span class="sxs-lookup"><span data-stu-id="fabfb-386">Object</span></span> | <span data-ttu-id="fabfb-387">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="fabfb-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="fabfb-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fabfb-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="fabfb-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fabfb-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="fabfb-394">Data</span><span class="sxs-lookup"><span data-stu-id="fabfb-394">Date</span></span> | <span data-ttu-id="fabfb-395">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fabfb-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="fabfb-396">Data</span><span class="sxs-lookup"><span data-stu-id="fabfb-396">Date</span></span> | <span data-ttu-id="fabfb-397">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fabfb-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="fabfb-398">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-398">String</span></span> | <span data-ttu-id="fabfb-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="fabfb-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="fabfb-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="fabfb-404">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-404">String</span></span> | <span data-ttu-id="fabfb-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="fabfb-407">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-407">String</span></span> | <span data-ttu-id="fabfb-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fabfb-410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-410">Requirements</span></span>

|<span data-ttu-id="fabfb-411">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-411">Requirement</span></span>| <span data-ttu-id="fabfb-412">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-414">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-414">1.0</span></span>|
|[<span data-ttu-id="fabfb-415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-416">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-418">Read</span><span class="sxs-lookup"><span data-stu-id="fabfb-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="fabfb-420">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="fabfb-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="fabfb-421">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="fabfb-422">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="fabfb-423">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="fabfb-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="fabfb-424">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fabfb-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-425">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-426">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="fabfb-426">All parameters are optional.</span></span>

|<span data-ttu-id="fabfb-427">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-427">Name</span></span>| <span data-ttu-id="fabfb-428">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-428">Type</span></span>| <span data-ttu-id="fabfb-429">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="fabfb-430">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-430">Object</span></span> | <span data-ttu-id="fabfb-431">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="fabfb-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fabfb-433">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="fabfb-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="fabfb-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="fabfb-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fabfb-436">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="fabfb-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="fabfb-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="fabfb-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="fabfb-439">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="fabfb-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="fabfb-440">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="fabfb-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="fabfb-441">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-441">String</span></span> | <span data-ttu-id="fabfb-442">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-442">A string containing the subject of the message.</span></span> <span data-ttu-id="fabfb-443">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fabfb-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="fabfb-444">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-444">String</span></span> | <span data-ttu-id="fabfb-445">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-445">The HTML body of the message.</span></span> <span data-ttu-id="fabfb-446">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fabfb-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="fabfb-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fabfb-448">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="fabfb-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="fabfb-449">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-449">String</span></span> | <span data-ttu-id="fabfb-p128">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="fabfb-452">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-452">String</span></span> | <span data-ttu-id="fabfb-453">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="fabfb-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="fabfb-454">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-454">String</span></span> | <span data-ttu-id="fabfb-p129">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="fabfb-457">Booliano</span><span class="sxs-lookup"><span data-stu-id="fabfb-457">Boolean</span></span> | <span data-ttu-id="fabfb-p130">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="fabfb-460">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fabfb-460">String</span></span> | <span data-ttu-id="fabfb-461">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="fabfb-462">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="fabfb-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="fabfb-463">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fabfb-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="fabfb-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-464">Requirements</span></span>

|<span data-ttu-id="fabfb-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-465">Requirement</span></span>| <span data-ttu-id="fabfb-466">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-468">1.6</span><span class="sxs-lookup"><span data-stu-id="fabfb-468">1.6</span></span> |
|[<span data-ttu-id="fabfb-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-470">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-472">Read</span><span class="sxs-lookup"><span data-stu-id="fabfb-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-473">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="fabfb-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fabfb-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="fabfb-475">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="fabfb-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="fabfb-p132">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-478">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="fabfb-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="fabfb-479">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="fabfb-479">**REST Tokens**</span></span>

<span data-ttu-id="fabfb-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="fabfb-483">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="fabfb-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="fabfb-484">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="fabfb-484">**EWS Tokens**</span></span>

<span data-ttu-id="fabfb-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="fabfb-487">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="fabfb-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-488">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-488">Parameters</span></span>

|<span data-ttu-id="fabfb-489">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-489">Name</span></span>| <span data-ttu-id="fabfb-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-490">Type</span></span>| <span data-ttu-id="fabfb-491">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-491">Attributes</span></span>| <span data-ttu-id="fabfb-492">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="fabfb-493">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-493">Object</span></span> | <span data-ttu-id="fabfb-494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-494">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-495">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fabfb-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="fabfb-496">Booliano</span><span class="sxs-lookup"><span data-stu-id="fabfb-496">Boolean</span></span> |  <span data-ttu-id="fabfb-497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-497">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fabfb-500">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-500">Object</span></span> |  <span data-ttu-id="fabfb-501">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-501">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-502">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fabfb-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="fabfb-503">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-503">function</span></span>||<span data-ttu-id="fabfb-p136">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-506">Requirements</span></span>

|<span data-ttu-id="fabfb-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-507">Requirement</span></span>| <span data-ttu-id="fabfb-508">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-510">1,5</span><span class="sxs-lookup"><span data-stu-id="fabfb-510">1.5</span></span> |
|[<span data-ttu-id="fabfb-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-512">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-514">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="fabfb-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="fabfb-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fabfb-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="fabfb-517">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="fabfb-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="fabfb-p137">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="fabfb-p138">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="fabfb-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="fabfb-523">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fabfb-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="fabfb-p139">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-526">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-526">Parameters</span></span>

|<span data-ttu-id="fabfb-527">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-527">Name</span></span>| <span data-ttu-id="fabfb-528">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-528">Type</span></span>| <span data-ttu-id="fabfb-529">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-529">Attributes</span></span>| <span data-ttu-id="fabfb-530">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fabfb-531">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-531">function</span></span>||<span data-ttu-id="fabfb-p140">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="fabfb-534">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-534">Object</span></span>| <span data-ttu-id="fabfb-535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-535">&lt;optional&gt;</span></span>|<span data-ttu-id="fabfb-536">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fabfb-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-537">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-537">Requirements</span></span>

|<span data-ttu-id="fabfb-538">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-538">Requirement</span></span>| <span data-ttu-id="fabfb-539">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-540">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-541">1.3</span><span class="sxs-lookup"><span data-stu-id="fabfb-541">1.3</span></span>|
|[<span data-ttu-id="fabfb-542">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-543">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-545">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="fabfb-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-546">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="fabfb-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fabfb-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="fabfb-548">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="fabfb-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="fabfb-549">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="fabfb-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-550">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-550">Parameters</span></span>

|<span data-ttu-id="fabfb-551">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-551">Name</span></span>| <span data-ttu-id="fabfb-552">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-552">Type</span></span>| <span data-ttu-id="fabfb-553">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-553">Attributes</span></span>| <span data-ttu-id="fabfb-554">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fabfb-555">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-555">function</span></span>||<span data-ttu-id="fabfb-556">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fabfb-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fabfb-557">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="fabfb-558">Object</span><span class="sxs-lookup"><span data-stu-id="fabfb-558">Object</span></span>| <span data-ttu-id="fabfb-559">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-559">&lt;optional&gt;</span></span>|<span data-ttu-id="fabfb-560">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fabfb-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-561">Requirements</span></span>

|<span data-ttu-id="fabfb-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-562">Requirement</span></span>| <span data-ttu-id="fabfb-563">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-565">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-565">1.0</span></span>|
|[<span data-ttu-id="fabfb-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-567">ReadItem</span></span>|
|[<span data-ttu-id="fabfb-568">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="fabfb-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-569">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="fabfb-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fabfb-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="fabfb-572">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="fabfb-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-573">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="fabfb-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="fabfb-574">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="fabfb-574">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="fabfb-575">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="fabfb-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="fabfb-576">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="fabfb-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="fabfb-577">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="fabfb-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="fabfb-578">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="fabfb-578">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="fabfb-579">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="fabfb-580">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="fabfb-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="fabfb-p142">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="fabfb-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="fabfb-583">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="fabfb-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="fabfb-584">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="fabfb-584">Version differences</span></span>

<span data-ttu-id="fabfb-585">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="fabfb-p143">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="fabfb-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-589">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-589">Parameters</span></span>

|<span data-ttu-id="fabfb-590">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-590">Name</span></span>| <span data-ttu-id="fabfb-591">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-591">Type</span></span>| <span data-ttu-id="fabfb-592">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-592">Attributes</span></span>| <span data-ttu-id="fabfb-593">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fabfb-594">String</span><span class="sxs-lookup"><span data-stu-id="fabfb-594">String</span></span>||<span data-ttu-id="fabfb-595">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="fabfb-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="fabfb-596">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-596">function</span></span>||<span data-ttu-id="fabfb-597">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fabfb-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fabfb-598">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="fabfb-599">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="fabfb-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="fabfb-600">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-600">Object</span></span>| <span data-ttu-id="fabfb-601">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-601">&lt;optional&gt;</span></span>|<span data-ttu-id="fabfb-602">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fabfb-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-603">Requirements</span></span>

|<span data-ttu-id="fabfb-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-604">Requirement</span></span>| <span data-ttu-id="fabfb-605">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-607">1.0</span><span class="sxs-lookup"><span data-stu-id="fabfb-607">1.0</span></span>|
|[<span data-ttu-id="fabfb-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="fabfb-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="fabfb-610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fabfb-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-611">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fabfb-612">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fabfb-612">Example</span></span>

<span data-ttu-id="fabfb-613">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="fabfb-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="fabfb-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fabfb-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="fabfb-615">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="fabfb-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="fabfb-616">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="fabfb-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fabfb-617">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="fabfb-617">Parameters</span></span>

| <span data-ttu-id="fabfb-618">Nome</span><span class="sxs-lookup"><span data-stu-id="fabfb-618">Name</span></span> | <span data-ttu-id="fabfb-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="fabfb-619">Type</span></span> | <span data-ttu-id="fabfb-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="fabfb-620">Attributes</span></span> | <span data-ttu-id="fabfb-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="fabfb-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fabfb-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fabfb-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fabfb-623">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="fabfb-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="fabfb-624">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-624">Object</span></span> | <span data-ttu-id="fabfb-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-625">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-626">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fabfb-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fabfb-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="fabfb-627">Object</span></span> | <span data-ttu-id="fabfb-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-628">&lt;optional&gt;</span></span> | <span data-ttu-id="fabfb-629">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fabfb-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fabfb-630">function</span><span class="sxs-lookup"><span data-stu-id="fabfb-630">function</span></span>| <span data-ttu-id="fabfb-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fabfb-631">&lt;optional&gt;</span></span>|<span data-ttu-id="fabfb-632">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fabfb-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fabfb-633">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fabfb-633">Requirements</span></span>

|<span data-ttu-id="fabfb-634">Requisito</span><span class="sxs-lookup"><span data-stu-id="fabfb-634">Requirement</span></span>| <span data-ttu-id="fabfb-635">Valor</span><span class="sxs-lookup"><span data-stu-id="fabfb-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="fabfb-636">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fabfb-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fabfb-637">1,5</span><span class="sxs-lookup"><span data-stu-id="fabfb-637">1.5</span></span> |
|[<span data-ttu-id="fabfb-638">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fabfb-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fabfb-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fabfb-639">ReadItem</span></span> |
|[<span data-ttu-id="fabfb-640">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="fabfb-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fabfb-641">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fabfb-641">Compose or Read</span></span>|
