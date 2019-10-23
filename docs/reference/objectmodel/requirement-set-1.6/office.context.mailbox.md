---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: b4bc64aa1ff836408a8b8b1efdaed7ddc8ce5725
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627080"
---
# <a name="mailbox"></a><span data-ttu-id="efc0e-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="efc0e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="efc0e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="efc0e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="efc0e-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="efc0e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="efc0e-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-105">Requirements</span></span>

|<span data-ttu-id="efc0e-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-106">Requirement</span></span>| <span data-ttu-id="efc0e-107">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-109">1.0</span></span>|
|[<span data-ttu-id="efc0e-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="efc0e-111">Restricted</span></span>|
|[<span data-ttu-id="efc0e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="efc0e-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="efc0e-114">Members and methods</span></span>

| <span data-ttu-id="efc0e-115">Membro</span><span class="sxs-lookup"><span data-stu-id="efc0e-115">Member</span></span> | <span data-ttu-id="efc0e-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="efc0e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="efc0e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="efc0e-118">Membro</span><span class="sxs-lookup"><span data-stu-id="efc0e-118">Member</span></span> |
| [<span data-ttu-id="efc0e-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="efc0e-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="efc0e-120">Membro</span><span class="sxs-lookup"><span data-stu-id="efc0e-120">Member</span></span> |
| [<span data-ttu-id="efc0e-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="efc0e-122">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-122">Method</span></span> |
| [<span data-ttu-id="efc0e-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="efc0e-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="efc0e-124">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-124">Method</span></span> |
| [<span data-ttu-id="efc0e-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="efc0e-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="efc0e-126">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-126">Method</span></span> |
| [<span data-ttu-id="efc0e-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="efc0e-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="efc0e-128">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-128">Method</span></span> |
| [<span data-ttu-id="efc0e-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="efc0e-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="efc0e-130">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-130">Method</span></span> |
| [<span data-ttu-id="efc0e-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="efc0e-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="efc0e-132">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-132">Method</span></span> |
| [<span data-ttu-id="efc0e-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="efc0e-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="efc0e-134">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-134">Method</span></span> |
| [<span data-ttu-id="efc0e-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="efc0e-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="efc0e-136">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-136">Method</span></span> |
| [<span data-ttu-id="efc0e-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="efc0e-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="efc0e-138">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-138">Method</span></span> |
| [<span data-ttu-id="efc0e-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="efc0e-140">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-140">Method</span></span> |
| [<span data-ttu-id="efc0e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="efc0e-142">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-142">Method</span></span> |
| [<span data-ttu-id="efc0e-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="efc0e-144">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-144">Method</span></span> |
| [<span data-ttu-id="efc0e-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="efc0e-146">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-146">Method</span></span> |
| [<span data-ttu-id="efc0e-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="efc0e-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="efc0e-148">Método</span><span class="sxs-lookup"><span data-stu-id="efc0e-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="efc0e-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="efc0e-149">Namespaces</span></span>

<span data-ttu-id="efc0e-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="efc0e-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="efc0e-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="efc0e-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="efc0e-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="efc0e-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="efc0e-153">Members</span><span class="sxs-lookup"><span data-stu-id="efc0e-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="efc0e-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="efc0e-154">ewsUrl: String</span></span>

<span data-ttu-id="efc0e-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-157">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="efc0e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="efc0e-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="efc0e-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="efc0e-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="efc0e-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-163">Type</span></span>

*   <span data-ttu-id="efc0e-164">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="efc0e-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-165">Requirements</span></span>

|<span data-ttu-id="efc0e-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-166">Requirement</span></span>| <span data-ttu-id="efc0e-167">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-169">1.0</span></span>|
|[<span data-ttu-id="efc0e-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-171">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="efc0e-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="efc0e-174">restUrl: String</span></span>

<span data-ttu-id="efc0e-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="efc0e-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="efc0e-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="efc0e-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="efc0e-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="efc0e-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="efc0e-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="efc0e-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-180">Type</span></span>

*   <span data-ttu-id="efc0e-181">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="efc0e-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-182">Requirements</span></span>

|<span data-ttu-id="efc0e-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-183">Requirement</span></span>| <span data-ttu-id="efc0e-184">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-186">1,5</span><span class="sxs-lookup"><span data-stu-id="efc0e-186">1.5</span></span> |
|[<span data-ttu-id="efc0e-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-188">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="efc0e-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="efc0e-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="efc0e-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="efc0e-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="efc0e-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="efc0e-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="efc0e-194">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="efc0e-195">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="efc0e-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-196">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-196">Parameters</span></span>

| <span data-ttu-id="efc0e-197">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-197">Name</span></span> | <span data-ttu-id="efc0e-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-198">Type</span></span> | <span data-ttu-id="efc0e-199">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-199">Attributes</span></span> | <span data-ttu-id="efc0e-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="efc0e-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="efc0e-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="efc0e-202">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="efc0e-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="efc0e-203">Função</span><span class="sxs-lookup"><span data-stu-id="efc0e-203">Function</span></span> || <span data-ttu-id="efc0e-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="efc0e-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-207">Object</span></span> | <span data-ttu-id="efc0e-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-208">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-209">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="efc0e-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="efc0e-210">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-210">Object</span></span> | <span data-ttu-id="efc0e-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-211">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-212">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="efc0e-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="efc0e-213">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-213">function</span></span>| <span data-ttu-id="efc0e-214">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-214">&lt;optional&gt;</span></span>|<span data-ttu-id="efc0e-215">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-216">Requirements</span></span>

|<span data-ttu-id="efc0e-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-217">Requirement</span></span>| <span data-ttu-id="efc0e-218">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-220">1,5</span><span class="sxs-lookup"><span data-stu-id="efc0e-220">1.5</span></span> |
|[<span data-ttu-id="efc0e-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-222">ReadItem</span></span> |
|[<span data-ttu-id="efc0e-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-225">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="efc0e-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="efc0e-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="efc0e-227">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="efc0e-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-228">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-231">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-231">Parameters</span></span>

|<span data-ttu-id="efc0e-232">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-232">Name</span></span>| <span data-ttu-id="efc0e-233">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-233">Type</span></span>| <span data-ttu-id="efc0e-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="efc0e-235">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-235">String</span></span>|<span data-ttu-id="efc0e-236">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="efc0e-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="efc0e-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="efc0e-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="efc0e-238">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-239">Requirements</span></span>

|<span data-ttu-id="efc0e-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-240">Requirement</span></span>| <span data-ttu-id="efc0e-241">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-243">1.3</span><span class="sxs-lookup"><span data-stu-id="efc0e-243">1.3</span></span>|
|[<span data-ttu-id="efc0e-244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-245">Restrito</span><span class="sxs-lookup"><span data-stu-id="efc0e-245">Restricted</span></span>|
|[<span data-ttu-id="efc0e-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="efc0e-248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="efc0e-248">Returns:</span></span>

<span data-ttu-id="efc0e-249">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="efc0e-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="efc0e-250">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-250">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="efc0e-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="efc0e-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="efc0e-252">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="efc0e-p108">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="efc0e-p109">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-258">Parameters</span></span>

|<span data-ttu-id="efc0e-259">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-259">Name</span></span>| <span data-ttu-id="efc0e-260">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-260">Type</span></span>| <span data-ttu-id="efc0e-261">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="efc0e-262">Date</span><span class="sxs-lookup"><span data-stu-id="efc0e-262">Date</span></span>|<span data-ttu-id="efc0e-263">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="efc0e-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-264">Requirements</span></span>

|<span data-ttu-id="efc0e-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-265">Requirement</span></span>| <span data-ttu-id="efc0e-266">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-268">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-268">1.0</span></span>|
|[<span data-ttu-id="efc0e-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-270">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="efc0e-273">Retorna:</span><span class="sxs-lookup"><span data-stu-id="efc0e-273">Returns:</span></span>

<span data-ttu-id="efc0e-274">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="efc0e-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="efc0e-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="efc0e-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="efc0e-276">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="efc0e-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-277">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-280">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-280">Parameters</span></span>

|<span data-ttu-id="efc0e-281">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-281">Name</span></span>| <span data-ttu-id="efc0e-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-282">Type</span></span>| <span data-ttu-id="efc0e-283">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="efc0e-284">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-284">String</span></span>|<span data-ttu-id="efc0e-285">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="efc0e-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="efc0e-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="efc0e-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="efc0e-287">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="efc0e-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-288">Requirements</span></span>

|<span data-ttu-id="efc0e-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-289">Requirement</span></span>| <span data-ttu-id="efc0e-290">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-292">1.3</span><span class="sxs-lookup"><span data-stu-id="efc0e-292">1.3</span></span>|
|[<span data-ttu-id="efc0e-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-294">Restrito</span><span class="sxs-lookup"><span data-stu-id="efc0e-294">Restricted</span></span>|
|[<span data-ttu-id="efc0e-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-296">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="efc0e-297">Retorna:</span><span class="sxs-lookup"><span data-stu-id="efc0e-297">Returns:</span></span>

<span data-ttu-id="efc0e-298">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="efc0e-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="efc0e-299">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-299">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="efc0e-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="efc0e-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="efc0e-301">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="efc0e-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="efc0e-302">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-303">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-303">Parameters</span></span>

|<span data-ttu-id="efc0e-304">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-304">Name</span></span>| <span data-ttu-id="efc0e-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-305">Type</span></span>| <span data-ttu-id="efc0e-306">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="efc0e-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="efc0e-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="efc0e-308">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="efc0e-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-309">Requirements</span></span>

|<span data-ttu-id="efc0e-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-310">Requirement</span></span>| <span data-ttu-id="efc0e-311">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-313">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-313">1.0</span></span>|
|[<span data-ttu-id="efc0e-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-315">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-317">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="efc0e-318">Retorna:</span><span class="sxs-lookup"><span data-stu-id="efc0e-318">Returns:</span></span>

<span data-ttu-id="efc0e-319">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="efc0e-319">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="efc0e-320">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="efc0e-320">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="efc0e-321">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-321">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="efc0e-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="efc0e-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="efc0e-323">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-324">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-325">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="efc0e-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="efc0e-p111">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="efc0e-328">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="efc0e-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="efc0e-329">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="efc0e-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-330">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-330">Parameters</span></span>

|<span data-ttu-id="efc0e-331">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-331">Name</span></span>| <span data-ttu-id="efc0e-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-332">Type</span></span>| <span data-ttu-id="efc0e-333">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="efc0e-334">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-334">String</span></span>|<span data-ttu-id="efc0e-335">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-336">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-336">Requirements</span></span>

|<span data-ttu-id="efc0e-337">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-337">Requirement</span></span>| <span data-ttu-id="efc0e-338">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-339">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-340">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-340">1.0</span></span>|
|[<span data-ttu-id="efc0e-341">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-342">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-343">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-344">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-345">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="efc0e-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="efc0e-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="efc0e-347">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-348">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-349">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="efc0e-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="efc0e-350">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="efc0e-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="efc0e-351">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="efc0e-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="efc0e-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-354">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-354">Parameters</span></span>

|<span data-ttu-id="efc0e-355">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-355">Name</span></span>| <span data-ttu-id="efc0e-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-356">Type</span></span>| <span data-ttu-id="efc0e-357">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="efc0e-358">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-358">String</span></span>|<span data-ttu-id="efc0e-359">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-360">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-360">Requirements</span></span>

|<span data-ttu-id="efc0e-361">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-361">Requirement</span></span>| <span data-ttu-id="efc0e-362">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-364">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-364">1.0</span></span>|
|[<span data-ttu-id="efc0e-365">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-366">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-367">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="efc0e-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-368">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-369">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="efc0e-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="efc0e-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="efc0e-371">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="efc0e-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-372">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="efc0e-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="efc0e-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="efc0e-p114">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="efc0e-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="efc0e-380">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="efc0e-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-381">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-382">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-382">All parameters are optional.</span></span>

|<span data-ttu-id="efc0e-383">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-383">Name</span></span>| <span data-ttu-id="efc0e-384">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-384">Type</span></span>| <span data-ttu-id="efc0e-385">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="efc0e-386">Object</span><span class="sxs-lookup"><span data-stu-id="efc0e-386">Object</span></span> | <span data-ttu-id="efc0e-387">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="efc0e-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="efc0e-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="efc0e-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="efc0e-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="efc0e-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="efc0e-394">Data</span><span class="sxs-lookup"><span data-stu-id="efc0e-394">Date</span></span> | <span data-ttu-id="efc0e-395">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="efc0e-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="efc0e-396">Data</span><span class="sxs-lookup"><span data-stu-id="efc0e-396">Date</span></span> | <span data-ttu-id="efc0e-397">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="efc0e-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="efc0e-398">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-398">String</span></span> | <span data-ttu-id="efc0e-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="efc0e-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="efc0e-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="efc0e-404">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-404">String</span></span> | <span data-ttu-id="efc0e-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="efc0e-407">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-407">String</span></span> | <span data-ttu-id="efc0e-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="efc0e-410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-410">Requirements</span></span>

|<span data-ttu-id="efc0e-411">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-411">Requirement</span></span>| <span data-ttu-id="efc0e-412">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-414">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-414">1.0</span></span>|
|[<span data-ttu-id="efc0e-415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-416">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-418">Read</span><span class="sxs-lookup"><span data-stu-id="efc0e-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="efc0e-420">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="efc0e-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="efc0e-421">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="efc0e-422">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="efc0e-423">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="efc0e-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="efc0e-424">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="efc0e-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-425">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-426">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-426">All parameters are optional.</span></span>

|<span data-ttu-id="efc0e-427">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-427">Name</span></span>| <span data-ttu-id="efc0e-428">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-428">Type</span></span>| <span data-ttu-id="efc0e-429">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="efc0e-430">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-430">Object</span></span> | <span data-ttu-id="efc0e-431">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="efc0e-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="efc0e-433">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="efc0e-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="efc0e-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="efc0e-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="efc0e-436">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="efc0e-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="efc0e-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="efc0e-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="efc0e-439">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="efc0e-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="efc0e-440">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="efc0e-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="efc0e-441">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-441">String</span></span> | <span data-ttu-id="efc0e-442">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-442">A string containing the subject of the message.</span></span> <span data-ttu-id="efc0e-443">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc0e-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="efc0e-444">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-444">String</span></span> | <span data-ttu-id="efc0e-445">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-445">The HTML body of the message.</span></span> <span data-ttu-id="efc0e-446">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="efc0e-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="efc0e-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="efc0e-448">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="efc0e-449">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-449">String</span></span> | <span data-ttu-id="efc0e-p128">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="efc0e-452">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-452">String</span></span> | <span data-ttu-id="efc0e-453">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="efc0e-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="efc0e-454">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-454">String</span></span> | <span data-ttu-id="efc0e-p129">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="efc0e-457">Booliano</span><span class="sxs-lookup"><span data-stu-id="efc0e-457">Boolean</span></span> | <span data-ttu-id="efc0e-p130">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="efc0e-460">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="efc0e-460">String</span></span> | <span data-ttu-id="efc0e-461">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="efc0e-462">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="efc0e-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="efc0e-463">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="efc0e-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="efc0e-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-464">Requirements</span></span>

|<span data-ttu-id="efc0e-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-465">Requirement</span></span>| <span data-ttu-id="efc0e-466">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-468">1.6</span><span class="sxs-lookup"><span data-stu-id="efc0e-468">1.6</span></span> |
|[<span data-ttu-id="efc0e-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-470">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-472">Read</span><span class="sxs-lookup"><span data-stu-id="efc0e-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="efc0e-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="efc0e-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="efc0e-475">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="efc0e-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="efc0e-p132">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-478">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="efc0e-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="efc0e-479">Chamar o `getCallbackTokenAsync` método no modo de leitura requer um nível mínimo de permissão de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-479">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="efc0e-480">A `getCallbackTokenAsync` chamada no modo de redação requer que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-480">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="efc0e-481">O [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) método requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-481">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="efc0e-482">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="efc0e-482">**REST Tokens**</span></span>

<span data-ttu-id="efc0e-p134">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="efc0e-486">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="efc0e-486">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="efc0e-487">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="efc0e-487">**EWS Tokens**</span></span>

<span data-ttu-id="efc0e-p135">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="efc0e-490">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="efc0e-490">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="efc0e-491">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="efc0e-491">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="efc0e-492">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-492">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="efc0e-493">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="efc0e-493">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-494">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-494">Parameters</span></span>

|<span data-ttu-id="efc0e-495">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-495">Name</span></span>| <span data-ttu-id="efc0e-496">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-496">Type</span></span>| <span data-ttu-id="efc0e-497">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-497">Attributes</span></span>| <span data-ttu-id="efc0e-498">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-498">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="efc0e-499">Object</span><span class="sxs-lookup"><span data-stu-id="efc0e-499">Object</span></span> | <span data-ttu-id="efc0e-500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-500">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-501">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="efc0e-501">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="efc0e-502">Booliano</span><span class="sxs-lookup"><span data-stu-id="efc0e-502">Boolean</span></span> |  <span data-ttu-id="efc0e-503">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-503">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-p137">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p137">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="efc0e-506">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-506">Object</span></span> |  <span data-ttu-id="efc0e-507">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-507">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-508">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="efc0e-508">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="efc0e-509">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-509">function</span></span>||<span data-ttu-id="efc0e-510">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-510">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="efc0e-511">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-511">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="efc0e-512">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-512">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="efc0e-513">Erros</span><span class="sxs-lookup"><span data-stu-id="efc0e-513">Errors</span></span>

|<span data-ttu-id="efc0e-514">Código de erro</span><span class="sxs-lookup"><span data-stu-id="efc0e-514">Error code</span></span>|<span data-ttu-id="efc0e-515">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-515">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="efc0e-516">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="efc0e-516">The request has failed.</span></span> <span data-ttu-id="efc0e-517">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="efc0e-517">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="efc0e-518">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="efc0e-518">The Exchange server returned an error.</span></span> <span data-ttu-id="efc0e-519">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="efc0e-519">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="efc0e-520">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="efc0e-520">The user is no longer connected to the network.</span></span> <span data-ttu-id="efc0e-521">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-521">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-522">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-522">Requirements</span></span>

|<span data-ttu-id="efc0e-523">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-523">Requirement</span></span>| <span data-ttu-id="efc0e-524">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-525">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-526">1,5</span><span class="sxs-lookup"><span data-stu-id="efc0e-526">1.5</span></span> |
|[<span data-ttu-id="efc0e-527">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-528">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-529">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-530">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="efc0e-530">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-531">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-531">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="efc0e-532">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="efc0e-532">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="efc0e-533">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="efc0e-533">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="efc0e-p141">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p141">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="efc0e-536">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="efc0e-536">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="efc0e-537">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-537">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="efc0e-538">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="efc0e-538">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="efc0e-539">Chamar o `getCallbackTokenAsync` método no modo de leitura requer um nível mínimo de permissão de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-539">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="efc0e-540">A `getCallbackTokenAsync` chamada no modo de redação requer que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-540">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="efc0e-541">O [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) método requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="efc0e-541">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-542">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-542">Parameters</span></span>

|<span data-ttu-id="efc0e-543">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-543">Name</span></span>| <span data-ttu-id="efc0e-544">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-544">Type</span></span>| <span data-ttu-id="efc0e-545">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-545">Attributes</span></span>| <span data-ttu-id="efc0e-546">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-546">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="efc0e-547">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-547">function</span></span>||<span data-ttu-id="efc0e-548">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-548">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="efc0e-549">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-549">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="efc0e-550">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-550">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="efc0e-551">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-551">Object</span></span>| <span data-ttu-id="efc0e-552">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-552">&lt;optional&gt;</span></span>|<span data-ttu-id="efc0e-553">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="efc0e-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="efc0e-554">Erros</span><span class="sxs-lookup"><span data-stu-id="efc0e-554">Errors</span></span>

|<span data-ttu-id="efc0e-555">Código de erro</span><span class="sxs-lookup"><span data-stu-id="efc0e-555">Error code</span></span>|<span data-ttu-id="efc0e-556">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-556">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="efc0e-557">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="efc0e-557">The request has failed.</span></span> <span data-ttu-id="efc0e-558">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="efc0e-558">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="efc0e-559">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="efc0e-559">The Exchange server returned an error.</span></span> <span data-ttu-id="efc0e-560">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="efc0e-560">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="efc0e-561">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="efc0e-561">The user is no longer connected to the network.</span></span> <span data-ttu-id="efc0e-562">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-562">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-563">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-563">Requirements</span></span>

|<span data-ttu-id="efc0e-564">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-564">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="efc0e-565">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-566">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-566">1.0</span></span> | <span data-ttu-id="efc0e-567">1.3</span><span class="sxs-lookup"><span data-stu-id="efc0e-567">1.3</span></span> |
|[<span data-ttu-id="efc0e-568">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-569">ReadItem</span></span> | <span data-ttu-id="efc0e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-570">ReadItem</span></span> |
|[<span data-ttu-id="efc0e-571">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-572">Read</span><span class="sxs-lookup"><span data-stu-id="efc0e-572">Read</span></span> | <span data-ttu-id="efc0e-573">Escrever</span><span class="sxs-lookup"><span data-stu-id="efc0e-573">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="efc0e-574">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-574">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="efc0e-575">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="efc0e-575">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="efc0e-576">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="efc0e-576">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="efc0e-577">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="efc0e-577">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-578">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-578">Parameters</span></span>

|<span data-ttu-id="efc0e-579">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-579">Name</span></span>| <span data-ttu-id="efc0e-580">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-580">Type</span></span>| <span data-ttu-id="efc0e-581">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-581">Attributes</span></span>| <span data-ttu-id="efc0e-582">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-582">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="efc0e-583">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-583">function</span></span>||<span data-ttu-id="efc0e-584">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="efc0e-585">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-585">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="efc0e-586">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="efc0e-586">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="efc0e-587">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-587">Object</span></span>| <span data-ttu-id="efc0e-588">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-588">&lt;optional&gt;</span></span>|<span data-ttu-id="efc0e-589">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="efc0e-589">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="efc0e-590">Erros</span><span class="sxs-lookup"><span data-stu-id="efc0e-590">Errors</span></span>

|<span data-ttu-id="efc0e-591">Código de erro</span><span class="sxs-lookup"><span data-stu-id="efc0e-591">Error code</span></span>|<span data-ttu-id="efc0e-592">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-592">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="efc0e-593">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="efc0e-593">The request has failed.</span></span> <span data-ttu-id="efc0e-594">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="efc0e-594">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="efc0e-595">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="efc0e-595">The Exchange server returned an error.</span></span> <span data-ttu-id="efc0e-596">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="efc0e-596">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="efc0e-597">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="efc0e-597">The user is no longer connected to the network.</span></span> <span data-ttu-id="efc0e-598">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="efc0e-598">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-599">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-599">Requirements</span></span>

|<span data-ttu-id="efc0e-600">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-600">Requirement</span></span>| <span data-ttu-id="efc0e-601">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-602">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-603">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-603">1.0</span></span>|
|[<span data-ttu-id="efc0e-604">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-605">ReadItem</span></span>|
|[<span data-ttu-id="efc0e-606">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="efc0e-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-607">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-607">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-608">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-608">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="efc0e-609">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="efc0e-609">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="efc0e-610">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="efc0e-610">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-611">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="efc0e-611">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="efc0e-612">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="efc0e-612">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="efc0e-613">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="efc0e-613">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="efc0e-614">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="efc0e-614">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="efc0e-615">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="efc0e-615">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="efc0e-616">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="efc0e-616">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="efc0e-617">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-617">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="efc0e-618">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="efc0e-618">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="efc0e-p151">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="efc0e-p151">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="efc0e-621">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="efc0e-621">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="efc0e-622">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="efc0e-622">Version differences</span></span>

<span data-ttu-id="efc0e-623">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-623">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="efc0e-p152">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="efc0e-p152">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-627">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-627">Parameters</span></span>

|<span data-ttu-id="efc0e-628">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-628">Name</span></span>| <span data-ttu-id="efc0e-629">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-629">Type</span></span>| <span data-ttu-id="efc0e-630">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-630">Attributes</span></span>| <span data-ttu-id="efc0e-631">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-631">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="efc0e-632">String</span><span class="sxs-lookup"><span data-stu-id="efc0e-632">String</span></span>||<span data-ttu-id="efc0e-633">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="efc0e-633">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="efc0e-634">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-634">function</span></span>||<span data-ttu-id="efc0e-635">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="efc0e-636">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-636">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="efc0e-637">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="efc0e-637">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="efc0e-638">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-638">Object</span></span>| <span data-ttu-id="efc0e-639">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-639">&lt;optional&gt;</span></span>|<span data-ttu-id="efc0e-640">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="efc0e-640">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-641">Requirements</span></span>

|<span data-ttu-id="efc0e-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-642">Requirement</span></span>| <span data-ttu-id="efc0e-643">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-644">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-645">1.0</span><span class="sxs-lookup"><span data-stu-id="efc0e-645">1.0</span></span>|
|[<span data-ttu-id="efc0e-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-647">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="efc0e-647">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="efc0e-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="efc0e-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-649">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-649">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="efc0e-650">Exemplo</span><span class="sxs-lookup"><span data-stu-id="efc0e-650">Example</span></span>

<span data-ttu-id="efc0e-651">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="efc0e-651">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="efc0e-652">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="efc0e-652">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="efc0e-653">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="efc0e-653">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="efc0e-654">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="efc0e-654">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="efc0e-655">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="efc0e-655">Parameters</span></span>

| <span data-ttu-id="efc0e-656">Nome</span><span class="sxs-lookup"><span data-stu-id="efc0e-656">Name</span></span> | <span data-ttu-id="efc0e-657">Tipo</span><span class="sxs-lookup"><span data-stu-id="efc0e-657">Type</span></span> | <span data-ttu-id="efc0e-658">Atributos</span><span class="sxs-lookup"><span data-stu-id="efc0e-658">Attributes</span></span> | <span data-ttu-id="efc0e-659">Descrição</span><span class="sxs-lookup"><span data-stu-id="efc0e-659">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="efc0e-660">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="efc0e-660">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="efc0e-661">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="efc0e-661">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="efc0e-662">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-662">Object</span></span> | <span data-ttu-id="efc0e-663">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-663">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-664">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="efc0e-664">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="efc0e-665">Objeto</span><span class="sxs-lookup"><span data-stu-id="efc0e-665">Object</span></span> | <span data-ttu-id="efc0e-666">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-666">&lt;optional&gt;</span></span> | <span data-ttu-id="efc0e-667">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="efc0e-667">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="efc0e-668">function</span><span class="sxs-lookup"><span data-stu-id="efc0e-668">function</span></span>| <span data-ttu-id="efc0e-669">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="efc0e-669">&lt;optional&gt;</span></span>|<span data-ttu-id="efc0e-670">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="efc0e-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="efc0e-671">Requisitos</span><span class="sxs-lookup"><span data-stu-id="efc0e-671">Requirements</span></span>

|<span data-ttu-id="efc0e-672">Requisito</span><span class="sxs-lookup"><span data-stu-id="efc0e-672">Requirement</span></span>| <span data-ttu-id="efc0e-673">Valor</span><span class="sxs-lookup"><span data-stu-id="efc0e-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="efc0e-674">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="efc0e-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="efc0e-675">1,5</span><span class="sxs-lookup"><span data-stu-id="efc0e-675">1.5</span></span> |
|[<span data-ttu-id="efc0e-676">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="efc0e-676">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="efc0e-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="efc0e-677">ReadItem</span></span> |
|[<span data-ttu-id="efc0e-678">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="efc0e-678">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="efc0e-679">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="efc0e-679">Compose or Read</span></span>|
