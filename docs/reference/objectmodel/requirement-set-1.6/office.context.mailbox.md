---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 09c3930daf6f26edbc38b01f515ee5b1830ce802
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629690"
---
# <a name="mailbox"></a><span data-ttu-id="c2ed5-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="c2ed5-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="c2ed5-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="c2ed5-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="c2ed5-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2ed5-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-105">Requirements</span></span>

|<span data-ttu-id="c2ed5-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-106">Requirement</span></span>| <span data-ttu-id="c2ed5-107">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-109">1.0</span></span>|
|[<span data-ttu-id="c2ed5-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-111">Restricted</span></span>|
|[<span data-ttu-id="c2ed5-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c2ed5-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-114">Members and methods</span></span>

| <span data-ttu-id="c2ed5-115">Membro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-115">Member</span></span> | <span data-ttu-id="c2ed5-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c2ed5-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c2ed5-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c2ed5-118">Membro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-118">Member</span></span> |
| [<span data-ttu-id="c2ed5-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="c2ed5-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c2ed5-120">Membro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-120">Member</span></span> |
| [<span data-ttu-id="c2ed5-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c2ed5-122">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-122">Method</span></span> |
| [<span data-ttu-id="c2ed5-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c2ed5-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c2ed5-124">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-124">Method</span></span> |
| [<span data-ttu-id="c2ed5-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c2ed5-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="c2ed5-126">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-126">Method</span></span> |
| [<span data-ttu-id="c2ed5-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c2ed5-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c2ed5-128">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-128">Method</span></span> |
| [<span data-ttu-id="c2ed5-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c2ed5-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c2ed5-130">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-130">Method</span></span> |
| [<span data-ttu-id="c2ed5-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c2ed5-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c2ed5-132">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-132">Method</span></span> |
| [<span data-ttu-id="c2ed5-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c2ed5-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c2ed5-134">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-134">Method</span></span> |
| [<span data-ttu-id="c2ed5-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c2ed5-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c2ed5-136">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-136">Method</span></span> |
| [<span data-ttu-id="c2ed5-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="c2ed5-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="c2ed5-138">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-138">Method</span></span> |
| [<span data-ttu-id="c2ed5-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c2ed5-140">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-140">Method</span></span> |
| [<span data-ttu-id="c2ed5-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c2ed5-142">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-142">Method</span></span> |
| [<span data-ttu-id="c2ed5-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c2ed5-144">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-144">Method</span></span> |
| [<span data-ttu-id="c2ed5-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c2ed5-146">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-146">Method</span></span> |
| [<span data-ttu-id="c2ed5-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c2ed5-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c2ed5-148">Método</span><span class="sxs-lookup"><span data-stu-id="c2ed5-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c2ed5-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="c2ed5-149">Namespaces</span></span>

<span data-ttu-id="c2ed5-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c2ed5-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c2ed5-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c2ed5-153">Members</span><span class="sxs-lookup"><span data-stu-id="c2ed5-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c2ed5-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-154">ewsUrl: String</span></span>

<span data-ttu-id="c2ed5-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-157">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c2ed5-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c2ed5-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c2ed5-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-163">Type</span></span>

*   <span data-ttu-id="c2ed5-164">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2ed5-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-165">Requirements</span></span>

|<span data-ttu-id="c2ed5-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-166">Requirement</span></span>| <span data-ttu-id="c2ed5-167">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-169">1.0</span></span>|
|[<span data-ttu-id="c2ed5-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-171">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="c2ed5-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-174">restUrl: String</span></span>

<span data-ttu-id="c2ed5-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c2ed5-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c2ed5-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-177">Type</span></span>

*   <span data-ttu-id="c2ed5-178">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2ed5-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-179">Requirements</span></span>

|<span data-ttu-id="c2ed5-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-180">Requirement</span></span>| <span data-ttu-id="c2ed5-181">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-183">1,5</span><span class="sxs-lookup"><span data-stu-id="c2ed5-183">1.5</span></span> |
|[<span data-ttu-id="c2ed5-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-185">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c2ed5-188">Métodos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c2ed5-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c2ed5-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c2ed5-190">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c2ed5-191">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-191">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="c2ed5-192">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-192">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-193">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-193">Parameters</span></span>

| <span data-ttu-id="c2ed5-194">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-194">Name</span></span> | <span data-ttu-id="c2ed5-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-195">Type</span></span> | <span data-ttu-id="c2ed5-196">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-196">Attributes</span></span> | <span data-ttu-id="c2ed5-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c2ed5-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c2ed5-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c2ed5-199">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c2ed5-200">Função</span><span class="sxs-lookup"><span data-stu-id="c2ed5-200">Function</span></span> || <span data-ttu-id="c2ed5-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c2ed5-204">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-204">Object</span></span> | <span data-ttu-id="c2ed5-205">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-205">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-206">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2ed5-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-207">Object</span></span> | <span data-ttu-id="c2ed5-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-208">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-209">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c2ed5-210">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-210">function</span></span>| <span data-ttu-id="c2ed5-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-211">&lt;optional&gt;</span></span>|<span data-ttu-id="c2ed5-212">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-213">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-213">Requirements</span></span>

|<span data-ttu-id="c2ed5-214">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-214">Requirement</span></span>| <span data-ttu-id="c2ed5-215">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-216">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-217">1,5</span><span class="sxs-lookup"><span data-stu-id="c2ed5-217">1.5</span></span> |
|[<span data-ttu-id="c2ed5-218">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-218">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-219">ReadItem</span></span> |
|[<span data-ttu-id="c2ed5-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-220">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-221">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-222">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c2ed5-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c2ed5-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c2ed5-224">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-225">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-225">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-228">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-228">Parameters</span></span>

|<span data-ttu-id="c2ed5-229">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-229">Name</span></span>| <span data-ttu-id="c2ed5-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-230">Type</span></span>| <span data-ttu-id="c2ed5-231">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2ed5-232">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-232">String</span></span>|<span data-ttu-id="c2ed5-233">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="c2ed5-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c2ed5-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c2ed5-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="c2ed5-235">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-236">Requirements</span></span>

|<span data-ttu-id="c2ed5-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-237">Requirement</span></span>| <span data-ttu-id="c2ed5-238">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-240">1.3</span><span class="sxs-lookup"><span data-stu-id="c2ed5-240">1.3</span></span>|
|[<span data-ttu-id="c2ed5-241">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-242">Restrito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-242">Restricted</span></span>|
|[<span data-ttu-id="c2ed5-243">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-244">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-244">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2ed5-245">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c2ed5-245">Returns:</span></span>

<span data-ttu-id="c2ed5-246">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c2ed5-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-247">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="c2ed5-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="c2ed5-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="c2ed5-249">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c2ed5-p107">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c2ed5-p108">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-255">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-255">Parameters</span></span>

|<span data-ttu-id="c2ed5-256">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-256">Name</span></span>| <span data-ttu-id="c2ed5-257">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-257">Type</span></span>| <span data-ttu-id="c2ed5-258">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c2ed5-259">Date</span><span class="sxs-lookup"><span data-stu-id="c2ed5-259">Date</span></span>|<span data-ttu-id="c2ed5-260">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="c2ed5-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-261">Requirements</span></span>

|<span data-ttu-id="c2ed5-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-262">Requirement</span></span>| <span data-ttu-id="c2ed5-263">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-264">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-265">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-265">1.0</span></span>|
|[<span data-ttu-id="c2ed5-266">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-266">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-267">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-268">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-268">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-269">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-269">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2ed5-270">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c2ed5-270">Returns:</span></span>

<span data-ttu-id="c2ed5-271">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c2ed5-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c2ed5-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c2ed5-273">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-274">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-274">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-277">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-277">Parameters</span></span>

|<span data-ttu-id="c2ed5-278">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-278">Name</span></span>| <span data-ttu-id="c2ed5-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-279">Type</span></span>| <span data-ttu-id="c2ed5-280">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2ed5-281">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-281">String</span></span>|<span data-ttu-id="c2ed5-282">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c2ed5-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c2ed5-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="c2ed5-284">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-285">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-285">Requirements</span></span>

|<span data-ttu-id="c2ed5-286">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-286">Requirement</span></span>| <span data-ttu-id="c2ed5-287">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-288">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-289">1.3</span><span class="sxs-lookup"><span data-stu-id="c2ed5-289">1.3</span></span>|
|[<span data-ttu-id="c2ed5-290">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-291">Restrito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-291">Restricted</span></span>|
|[<span data-ttu-id="c2ed5-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-293">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2ed5-294">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c2ed5-294">Returns:</span></span>

<span data-ttu-id="c2ed5-295">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c2ed5-296">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-296">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c2ed5-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c2ed5-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c2ed5-298">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c2ed5-299">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-300">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-300">Parameters</span></span>

|<span data-ttu-id="c2ed5-301">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-301">Name</span></span>| <span data-ttu-id="c2ed5-302">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-302">Type</span></span>| <span data-ttu-id="c2ed5-303">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c2ed5-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c2ed5-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="c2ed5-305">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-306">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-306">Requirements</span></span>

|<span data-ttu-id="c2ed5-307">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-307">Requirement</span></span>| <span data-ttu-id="c2ed5-308">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-309">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-310">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-310">1.0</span></span>|
|[<span data-ttu-id="c2ed5-311">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-312">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-314">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2ed5-315">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c2ed5-315">Returns:</span></span>

<span data-ttu-id="c2ed5-316">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-316">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="c2ed5-317">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="c2ed5-317">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="c2ed5-318">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-318">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="c2ed5-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c2ed5-320">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-321">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-322">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c2ed5-p110">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c2ed5-325">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-325">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c2ed5-326">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-327">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-327">Parameters</span></span>

|<span data-ttu-id="c2ed5-328">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-328">Name</span></span>| <span data-ttu-id="c2ed5-329">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-329">Type</span></span>| <span data-ttu-id="c2ed5-330">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2ed5-331">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-331">String</span></span>|<span data-ttu-id="c2ed5-332">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-333">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-333">Requirements</span></span>

|<span data-ttu-id="c2ed5-334">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-334">Requirement</span></span>| <span data-ttu-id="c2ed5-335">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-336">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-337">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-337">1.0</span></span>|
|[<span data-ttu-id="c2ed5-338">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-339">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-340">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-341">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-341">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-342">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-342">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="c2ed5-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c2ed5-344">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-345">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-345">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-346">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c2ed5-347">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-347">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c2ed5-348">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c2ed5-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-351">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-351">Parameters</span></span>

|<span data-ttu-id="c2ed5-352">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-352">Name</span></span>| <span data-ttu-id="c2ed5-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-353">Type</span></span>| <span data-ttu-id="c2ed5-354">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2ed5-355">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-355">String</span></span>|<span data-ttu-id="c2ed5-356">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-357">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-357">Requirements</span></span>

|<span data-ttu-id="c2ed5-358">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-358">Requirement</span></span>| <span data-ttu-id="c2ed5-359">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-360">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-361">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-361">1.0</span></span>|
|[<span data-ttu-id="c2ed5-362">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-363">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-364">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c2ed5-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-365">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-365">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-366">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-366">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c2ed5-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c2ed5-368">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-369">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-369">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c2ed5-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c2ed5-p113">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c2ed5-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c2ed5-377">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-378">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-378">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-379">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-379">All parameters are optional.</span></span>

|<span data-ttu-id="c2ed5-380">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-380">Name</span></span>| <span data-ttu-id="c2ed5-381">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-381">Type</span></span>| <span data-ttu-id="c2ed5-382">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c2ed5-383">Object</span><span class="sxs-lookup"><span data-stu-id="c2ed5-383">Object</span></span> | <span data-ttu-id="c2ed5-384">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c2ed5-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c2ed5-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c2ed5-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c2ed5-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c2ed5-391">Data</span><span class="sxs-lookup"><span data-stu-id="c2ed5-391">Date</span></span> | <span data-ttu-id="c2ed5-392">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c2ed5-393">Data</span><span class="sxs-lookup"><span data-stu-id="c2ed5-393">Date</span></span> | <span data-ttu-id="c2ed5-394">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c2ed5-395">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-395">String</span></span> | <span data-ttu-id="c2ed5-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c2ed5-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c2ed5-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c2ed5-401">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-401">String</span></span> | <span data-ttu-id="c2ed5-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c2ed5-404">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-404">String</span></span> | <span data-ttu-id="c2ed5-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c2ed5-407">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-407">Requirements</span></span>

|<span data-ttu-id="c2ed5-408">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-408">Requirement</span></span>| <span data-ttu-id="c2ed5-409">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-410">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-411">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-411">1.0</span></span>|
|[<span data-ttu-id="c2ed5-412">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-413">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-414">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-415">Read</span><span class="sxs-lookup"><span data-stu-id="c2ed5-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-416">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="c2ed5-417">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="c2ed5-418">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="c2ed5-419">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="c2ed5-420">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c2ed5-421">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-422">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-422">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-423">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-423">All parameters are optional.</span></span>

|<span data-ttu-id="c2ed5-424">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-424">Name</span></span>| <span data-ttu-id="c2ed5-425">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-425">Type</span></span>| <span data-ttu-id="c2ed5-426">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c2ed5-427">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-427">Object</span></span> | <span data-ttu-id="c2ed5-428">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="c2ed5-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c2ed5-430">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="c2ed5-431">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="c2ed5-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c2ed5-433">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="c2ed5-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="c2ed5-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c2ed5-436">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="c2ed5-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c2ed5-438">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-438">String</span></span> | <span data-ttu-id="c2ed5-439">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-439">A string containing the subject of the message.</span></span> <span data-ttu-id="c2ed5-440">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="c2ed5-441">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-441">String</span></span> | <span data-ttu-id="c2ed5-442">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-442">The HTML body of the message.</span></span> <span data-ttu-id="c2ed5-443">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="c2ed5-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c2ed5-445">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="c2ed5-446">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-446">String</span></span> | <span data-ttu-id="c2ed5-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="c2ed5-449">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-449">String</span></span> | <span data-ttu-id="c2ed5-450">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="c2ed5-451">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-451">String</span></span> | <span data-ttu-id="c2ed5-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="c2ed5-454">Booliano</span><span class="sxs-lookup"><span data-stu-id="c2ed5-454">Boolean</span></span> | <span data-ttu-id="c2ed5-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="c2ed5-457">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c2ed5-457">String</span></span> | <span data-ttu-id="c2ed5-458">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="c2ed5-459">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="c2ed5-460">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="c2ed5-461">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-461">Requirements</span></span>

|<span data-ttu-id="c2ed5-462">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-462">Requirement</span></span>| <span data-ttu-id="c2ed5-463">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-464">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-465">1.6</span><span class="sxs-lookup"><span data-stu-id="c2ed5-465">1.6</span></span> |
|[<span data-ttu-id="c2ed5-466">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-466">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-467">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-468">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-468">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-469">Read</span><span class="sxs-lookup"><span data-stu-id="c2ed5-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-470">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-470">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c2ed5-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c2ed5-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c2ed5-472">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c2ed5-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-475">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="c2ed5-476">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-476">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="c2ed5-477">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-477">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="c2ed5-478">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-478">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="c2ed5-479">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="c2ed5-479">**REST Tokens**</span></span>

<span data-ttu-id="c2ed5-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c2ed5-483">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c2ed5-484">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="c2ed5-484">**EWS Tokens**</span></span>

<span data-ttu-id="c2ed5-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c2ed5-487">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="c2ed5-488">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-488">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="c2ed5-489">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-489">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="c2ed5-490">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-490">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-491">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-491">Parameters</span></span>

|<span data-ttu-id="c2ed5-492">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-492">Name</span></span>| <span data-ttu-id="c2ed5-493">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-493">Type</span></span>| <span data-ttu-id="c2ed5-494">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-494">Attributes</span></span>| <span data-ttu-id="c2ed5-495">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-495">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c2ed5-496">Object</span><span class="sxs-lookup"><span data-stu-id="c2ed5-496">Object</span></span> | <span data-ttu-id="c2ed5-497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-497">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-498">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-498">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c2ed5-499">Booliano</span><span class="sxs-lookup"><span data-stu-id="c2ed5-499">Boolean</span></span> |  <span data-ttu-id="c2ed5-500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-500">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-p136">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2ed5-503">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-503">Object</span></span> |  <span data-ttu-id="c2ed5-504">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-504">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-505">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-505">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c2ed5-506">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-506">function</span></span>||<span data-ttu-id="c2ed5-507">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-507">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2ed5-508">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-508">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c2ed5-509">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-509">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c2ed5-510">Erros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-510">Errors</span></span>

|<span data-ttu-id="c2ed5-511">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-511">Error code</span></span>|<span data-ttu-id="c2ed5-512">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-512">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c2ed5-513">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-513">The request has failed.</span></span> <span data-ttu-id="c2ed5-514">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-514">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c2ed5-515">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-515">The Exchange server returned an error.</span></span> <span data-ttu-id="c2ed5-516">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-516">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c2ed5-517">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-517">The user is no longer connected to the network.</span></span> <span data-ttu-id="c2ed5-518">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-518">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-519">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-519">Requirements</span></span>

|<span data-ttu-id="c2ed5-520">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-520">Requirement</span></span>| <span data-ttu-id="c2ed5-521">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-522">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-523">1,5</span><span class="sxs-lookup"><span data-stu-id="c2ed5-523">1.5</span></span> |
|[<span data-ttu-id="c2ed5-524">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-525">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-526">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-527">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="c2ed5-527">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-528">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-528">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c2ed5-529">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2ed5-529">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c2ed5-530">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-530">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c2ed5-p140">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c2ed5-533">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-533">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="c2ed5-534">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-534">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="c2ed5-535">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-535">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c2ed5-536">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-536">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="c2ed5-537">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-537">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="c2ed5-538">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-538">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-539">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-539">Parameters</span></span>

|<span data-ttu-id="c2ed5-540">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-540">Name</span></span>| <span data-ttu-id="c2ed5-541">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-541">Type</span></span>| <span data-ttu-id="c2ed5-542">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-542">Attributes</span></span>| <span data-ttu-id="c2ed5-543">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-543">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c2ed5-544">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-544">function</span></span>||<span data-ttu-id="c2ed5-545">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2ed5-546">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-546">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c2ed5-547">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-547">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="c2ed5-548">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-548">Object</span></span>| <span data-ttu-id="c2ed5-549">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-549">&lt;optional&gt;</span></span>|<span data-ttu-id="c2ed5-550">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-550">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c2ed5-551">Erros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-551">Errors</span></span>

|<span data-ttu-id="c2ed5-552">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-552">Error code</span></span>|<span data-ttu-id="c2ed5-553">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-553">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c2ed5-554">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-554">The request has failed.</span></span> <span data-ttu-id="c2ed5-555">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-555">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c2ed5-556">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-556">The Exchange server returned an error.</span></span> <span data-ttu-id="c2ed5-557">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-557">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c2ed5-558">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-558">The user is no longer connected to the network.</span></span> <span data-ttu-id="c2ed5-559">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-559">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-560">Requirements</span></span>

|<span data-ttu-id="c2ed5-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-561">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c2ed5-562">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-563">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-563">1.0</span></span> | <span data-ttu-id="c2ed5-564">1.3</span><span class="sxs-lookup"><span data-stu-id="c2ed5-564">1.3</span></span> |
|[<span data-ttu-id="c2ed5-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-566">ReadItem</span></span> | <span data-ttu-id="c2ed5-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-567">ReadItem</span></span> |
|[<span data-ttu-id="c2ed5-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-569">Read</span><span class="sxs-lookup"><span data-stu-id="c2ed5-569">Read</span></span> | <span data-ttu-id="c2ed5-570">Escrever</span><span class="sxs-lookup"><span data-stu-id="c2ed5-570">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="c2ed5-571">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-571">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c2ed5-572">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2ed5-572">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c2ed5-573">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-573">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c2ed5-574">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-574">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-575">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-575">Parameters</span></span>

|<span data-ttu-id="c2ed5-576">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-576">Name</span></span>| <span data-ttu-id="c2ed5-577">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-577">Type</span></span>| <span data-ttu-id="c2ed5-578">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-578">Attributes</span></span>| <span data-ttu-id="c2ed5-579">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-579">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c2ed5-580">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-580">function</span></span>||<span data-ttu-id="c2ed5-581">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2ed5-582">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-582">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c2ed5-583">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-583">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="c2ed5-584">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-584">Object</span></span>| <span data-ttu-id="c2ed5-585">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-585">&lt;optional&gt;</span></span>|<span data-ttu-id="c2ed5-586">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-586">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c2ed5-587">Erros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-587">Errors</span></span>

|<span data-ttu-id="c2ed5-588">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c2ed5-588">Error code</span></span>|<span data-ttu-id="c2ed5-589">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-589">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c2ed5-590">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-590">The request has failed.</span></span> <span data-ttu-id="c2ed5-591">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-591">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c2ed5-592">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-592">The Exchange server returned an error.</span></span> <span data-ttu-id="c2ed5-593">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-593">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c2ed5-594">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-594">The user is no longer connected to the network.</span></span> <span data-ttu-id="c2ed5-595">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-595">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-596">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-596">Requirements</span></span>

|<span data-ttu-id="c2ed5-597">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-597">Requirement</span></span>| <span data-ttu-id="c2ed5-598">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-599">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-600">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-600">1.0</span></span>|
|[<span data-ttu-id="c2ed5-601">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-602">ReadItem</span></span>|
|[<span data-ttu-id="c2ed5-603">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c2ed5-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-604">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-604">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-605">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-605">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c2ed5-606">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2ed5-606">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c2ed5-607">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-607">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-608">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-608">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c2ed5-609">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="c2ed5-609">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="c2ed5-610">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="c2ed5-610">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c2ed5-611">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-611">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c2ed5-612">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-612">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="c2ed5-613">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-613">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c2ed5-614">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-614">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c2ed5-615">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-615">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c2ed5-p150">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c2ed5-618">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-618">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c2ed5-619">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="c2ed5-619">Version differences</span></span>

<span data-ttu-id="c2ed5-620">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-620">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c2ed5-p151">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-624">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-624">Parameters</span></span>

|<span data-ttu-id="c2ed5-625">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-625">Name</span></span>| <span data-ttu-id="c2ed5-626">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-626">Type</span></span>| <span data-ttu-id="c2ed5-627">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-627">Attributes</span></span>| <span data-ttu-id="c2ed5-628">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-628">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c2ed5-629">String</span><span class="sxs-lookup"><span data-stu-id="c2ed5-629">String</span></span>||<span data-ttu-id="c2ed5-630">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-630">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c2ed5-631">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-631">function</span></span>||<span data-ttu-id="c2ed5-632">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2ed5-633">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-633">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="c2ed5-634">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-634">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c2ed5-635">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-635">Object</span></span>| <span data-ttu-id="c2ed5-636">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-636">&lt;optional&gt;</span></span>|<span data-ttu-id="c2ed5-637">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-637">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-638">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-638">Requirements</span></span>

|<span data-ttu-id="c2ed5-639">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-639">Requirement</span></span>| <span data-ttu-id="c2ed5-640">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-640">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-641">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-641">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-642">1.0</span><span class="sxs-lookup"><span data-stu-id="c2ed5-642">1.0</span></span>|
|[<span data-ttu-id="c2ed5-643">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-643">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-644">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c2ed5-644">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c2ed5-645">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c2ed5-645">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-646">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-646">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2ed5-647">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-647">Example</span></span>

<span data-ttu-id="c2ed5-648">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-648">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c2ed5-649">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c2ed5-649">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c2ed5-650">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-650">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c2ed5-651">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-651">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2ed5-652">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c2ed5-652">Parameters</span></span>

| <span data-ttu-id="c2ed5-653">Nome</span><span class="sxs-lookup"><span data-stu-id="c2ed5-653">Name</span></span> | <span data-ttu-id="c2ed5-654">Tipo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-654">Type</span></span> | <span data-ttu-id="c2ed5-655">Atributos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-655">Attributes</span></span> | <span data-ttu-id="c2ed5-656">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2ed5-656">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c2ed5-657">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c2ed5-657">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c2ed5-658">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-658">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c2ed5-659">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-659">Object</span></span> | <span data-ttu-id="c2ed5-660">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-660">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-661">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-661">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2ed5-662">Objeto</span><span class="sxs-lookup"><span data-stu-id="c2ed5-662">Object</span></span> | <span data-ttu-id="c2ed5-663">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-663">&lt;optional&gt;</span></span> | <span data-ttu-id="c2ed5-664">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c2ed5-664">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c2ed5-665">function</span><span class="sxs-lookup"><span data-stu-id="c2ed5-665">function</span></span>| <span data-ttu-id="c2ed5-666">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c2ed5-666">&lt;optional&gt;</span></span>|<span data-ttu-id="c2ed5-667">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c2ed5-667">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2ed5-668">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c2ed5-668">Requirements</span></span>

|<span data-ttu-id="c2ed5-669">Requisito</span><span class="sxs-lookup"><span data-stu-id="c2ed5-669">Requirement</span></span>| <span data-ttu-id="c2ed5-670">Valor</span><span class="sxs-lookup"><span data-stu-id="c2ed5-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2ed5-671">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c2ed5-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2ed5-672">1,5</span><span class="sxs-lookup"><span data-stu-id="c2ed5-672">1.5</span></span> |
|[<span data-ttu-id="c2ed5-673">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c2ed5-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2ed5-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2ed5-674">ReadItem</span></span> |
|[<span data-ttu-id="c2ed5-675">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c2ed5-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c2ed5-676">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c2ed5-676">Compose or Read</span></span>|
