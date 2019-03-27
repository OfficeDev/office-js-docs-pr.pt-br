---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9b91a61d301434886723a55eca9608f004f598eb
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871931"
---
# <a name="mailbox"></a><span data-ttu-id="d9852-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="d9852-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d9852-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d9852-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d9852-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="d9852-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9852-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-105">Requirements</span></span>

|<span data-ttu-id="d9852-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-106">Requirement</span></span>| <span data-ttu-id="d9852-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-109">1.0</span></span>|
|[<span data-ttu-id="d9852-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="d9852-111">Restricted</span></span>|
|[<span data-ttu-id="d9852-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d9852-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="d9852-114">Members and methods</span></span>

| <span data-ttu-id="d9852-115">Membro</span><span class="sxs-lookup"><span data-stu-id="d9852-115">Member</span></span> | <span data-ttu-id="d9852-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d9852-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d9852-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d9852-118">Member</span><span class="sxs-lookup"><span data-stu-id="d9852-118">Member</span></span> |
| [<span data-ttu-id="d9852-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d9852-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d9852-120">Membro</span><span class="sxs-lookup"><span data-stu-id="d9852-120">Member</span></span> |
| [<span data-ttu-id="d9852-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d9852-122">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-122">Method</span></span> |
| [<span data-ttu-id="d9852-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d9852-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d9852-124">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-124">Method</span></span> |
| [<span data-ttu-id="d9852-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d9852-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d9852-126">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-126">Method</span></span> |
| [<span data-ttu-id="d9852-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d9852-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d9852-128">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-128">Method</span></span> |
| [<span data-ttu-id="d9852-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d9852-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d9852-130">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-130">Method</span></span> |
| [<span data-ttu-id="d9852-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d9852-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d9852-132">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-132">Method</span></span> |
| [<span data-ttu-id="d9852-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d9852-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d9852-134">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-134">Method</span></span> |
| [<span data-ttu-id="d9852-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d9852-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d9852-136">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-136">Method</span></span> |
| [<span data-ttu-id="d9852-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="d9852-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d9852-138">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-138">Method</span></span> |
| [<span data-ttu-id="d9852-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d9852-140">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-140">Method</span></span> |
| [<span data-ttu-id="d9852-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d9852-142">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-142">Method</span></span> |
| [<span data-ttu-id="d9852-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d9852-144">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-144">Method</span></span> |
| [<span data-ttu-id="d9852-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d9852-146">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-146">Method</span></span> |
| [<span data-ttu-id="d9852-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d9852-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d9852-148">Método</span><span class="sxs-lookup"><span data-stu-id="d9852-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d9852-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d9852-149">Namespaces</span></span>

<span data-ttu-id="d9852-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d9852-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d9852-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d9852-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d9852-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d9852-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d9852-153">Membros</span><span class="sxs-lookup"><span data-stu-id="d9852-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d9852-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="d9852-154">ewsUrl :String</span></span>

<span data-ttu-id="d9852-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="d9852-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-157">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d9852-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d9852-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d9852-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d9852-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9852-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-163">Type</span></span>

*   <span data-ttu-id="d9852-164">String</span><span class="sxs-lookup"><span data-stu-id="d9852-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9852-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-165">Requirements</span></span>

|<span data-ttu-id="d9852-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-166">Requirement</span></span>| <span data-ttu-id="d9852-167">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-169">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-169">1.0</span></span>|
|[<span data-ttu-id="d9852-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-171">ReadItem</span></span>|
|[<span data-ttu-id="d9852-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="d9852-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="d9852-174">restUrl :String</span></span>

<span data-ttu-id="d9852-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="d9852-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d9852-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d9852-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d9852-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d9852-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d9852-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9852-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-180">Type</span></span>

*   <span data-ttu-id="d9852-181">String</span><span class="sxs-lookup"><span data-stu-id="d9852-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9852-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-182">Requirements</span></span>

|<span data-ttu-id="d9852-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-183">Requirement</span></span>| <span data-ttu-id="d9852-184">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-186">1,5</span><span class="sxs-lookup"><span data-stu-id="d9852-186">1.5</span></span> |
|[<span data-ttu-id="d9852-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-188">ReadItem</span></span>|
|[<span data-ttu-id="d9852-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d9852-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="d9852-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d9852-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9852-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d9852-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="d9852-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d9852-194">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="d9852-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="d9852-195">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="d9852-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-196">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-196">Parameters</span></span>

| <span data-ttu-id="d9852-197">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-197">Name</span></span> | <span data-ttu-id="d9852-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-198">Type</span></span> | <span data-ttu-id="d9852-199">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-199">Attributes</span></span> | <span data-ttu-id="d9852-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d9852-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d9852-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d9852-202">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="d9852-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d9852-203">Função</span><span class="sxs-lookup"><span data-stu-id="d9852-203">Function</span></span> || <span data-ttu-id="d9852-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d9852-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-207">Object</span></span> | <span data-ttu-id="d9852-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-208">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-209">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d9852-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d9852-210">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-210">Object</span></span> | <span data-ttu-id="d9852-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-211">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-212">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d9852-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d9852-213">function</span><span class="sxs-lookup"><span data-stu-id="d9852-213">function</span></span>| <span data-ttu-id="d9852-214">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-214">&lt;optional&gt;</span></span>|<span data-ttu-id="d9852-215">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9852-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-216">Requirements</span></span>

|<span data-ttu-id="d9852-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-217">Requirement</span></span>| <span data-ttu-id="d9852-218">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-220">1,5</span><span class="sxs-lookup"><span data-stu-id="d9852-220">1.5</span></span> |
|[<span data-ttu-id="d9852-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-222">ReadItem</span></span> |
|[<span data-ttu-id="d9852-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d9852-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d9852-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d9852-227">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="d9852-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-228">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="d9852-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-231">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-231">Parameters</span></span>

|<span data-ttu-id="d9852-232">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-232">Name</span></span>| <span data-ttu-id="d9852-233">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-233">Type</span></span>| <span data-ttu-id="d9852-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d9852-235">String</span><span class="sxs-lookup"><span data-stu-id="d9852-235">String</span></span>|<span data-ttu-id="d9852-236">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="d9852-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d9852-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d9852-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="d9852-238">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="d9852-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-239">Requirements</span></span>

|<span data-ttu-id="d9852-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-240">Requirement</span></span>| <span data-ttu-id="d9852-241">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-243">1.3</span><span class="sxs-lookup"><span data-stu-id="d9852-243">1.3</span></span>|
|[<span data-ttu-id="d9852-244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-245">Restrito</span><span class="sxs-lookup"><span data-stu-id="d9852-245">Restricted</span></span>|
|[<span data-ttu-id="d9852-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9852-248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d9852-248">Returns:</span></span>

<span data-ttu-id="d9852-249">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="d9852-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d9852-250">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="d9852-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d9852-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="d9852-252">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="d9852-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d9852-p108">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="d9852-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d9852-p109">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="d9852-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-258">Parameters</span></span>

|<span data-ttu-id="d9852-259">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-259">Name</span></span>| <span data-ttu-id="d9852-260">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-260">Type</span></span>| <span data-ttu-id="d9852-261">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d9852-262">Data</span><span class="sxs-lookup"><span data-stu-id="d9852-262">Date</span></span>|<span data-ttu-id="d9852-263">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="d9852-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-264">Requirements</span></span>

|<span data-ttu-id="d9852-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-265">Requirement</span></span>| <span data-ttu-id="d9852-266">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-268">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-268">1.0</span></span>|
|[<span data-ttu-id="d9852-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-270">ReadItem</span></span>|
|[<span data-ttu-id="d9852-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9852-273">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d9852-273">Returns:</span></span>

<span data-ttu-id="d9852-274">Tipo: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d9852-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d9852-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d9852-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d9852-276">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="d9852-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-277">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="d9852-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-280">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-280">Parameters</span></span>

|<span data-ttu-id="d9852-281">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-281">Name</span></span>| <span data-ttu-id="d9852-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-282">Type</span></span>| <span data-ttu-id="d9852-283">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d9852-284">String</span><span class="sxs-lookup"><span data-stu-id="d9852-284">String</span></span>|<span data-ttu-id="d9852-285">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="d9852-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d9852-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d9852-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="d9852-287">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="d9852-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-288">Requirements</span></span>

|<span data-ttu-id="d9852-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-289">Requirement</span></span>| <span data-ttu-id="d9852-290">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-292">1.3</span><span class="sxs-lookup"><span data-stu-id="d9852-292">1.3</span></span>|
|[<span data-ttu-id="d9852-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-294">Restrito</span><span class="sxs-lookup"><span data-stu-id="d9852-294">Restricted</span></span>|
|[<span data-ttu-id="d9852-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-296">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9852-297">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d9852-297">Returns:</span></span>

<span data-ttu-id="d9852-298">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="d9852-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d9852-299">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d9852-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d9852-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d9852-301">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="d9852-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d9852-302">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="d9852-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-303">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-303">Parameters</span></span>

|<span data-ttu-id="d9852-304">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-304">Name</span></span>| <span data-ttu-id="d9852-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-305">Type</span></span>| <span data-ttu-id="d9852-306">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d9852-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d9852-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="d9852-308">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="d9852-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-309">Requirements</span></span>

|<span data-ttu-id="d9852-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-310">Requirement</span></span>| <span data-ttu-id="d9852-311">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-313">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-313">1.0</span></span>|
|[<span data-ttu-id="d9852-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-315">ReadItem</span></span>|
|[<span data-ttu-id="d9852-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-317">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9852-318">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d9852-318">Returns:</span></span>

<span data-ttu-id="d9852-319">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="d9852-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="d9852-320">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="d9852-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9852-321">Date</span><span class="sxs-lookup"><span data-stu-id="d9852-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="d9852-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d9852-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d9852-323">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="d9852-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-324">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-325">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="d9852-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d9852-p111">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="d9852-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d9852-328">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d9852-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d9852-329">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="d9852-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-330">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-330">Parameters</span></span>

|<span data-ttu-id="d9852-331">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-331">Name</span></span>| <span data-ttu-id="d9852-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-332">Type</span></span>| <span data-ttu-id="d9852-333">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d9852-334">String</span><span class="sxs-lookup"><span data-stu-id="d9852-334">String</span></span>|<span data-ttu-id="d9852-335">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="d9852-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-336">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-336">Requirements</span></span>

|<span data-ttu-id="d9852-337">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-337">Requirement</span></span>| <span data-ttu-id="d9852-338">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-339">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-340">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-340">1.0</span></span>|
|[<span data-ttu-id="d9852-341">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-342">ReadItem</span></span>|
|[<span data-ttu-id="d9852-343">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d9852-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-344">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-345">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="d9852-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d9852-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d9852-347">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="d9852-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-348">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-349">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="d9852-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d9852-350">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d9852-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d9852-351">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="d9852-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d9852-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="d9852-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-354">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-354">Parameters</span></span>

|<span data-ttu-id="d9852-355">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-355">Name</span></span>| <span data-ttu-id="d9852-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-356">Type</span></span>| <span data-ttu-id="d9852-357">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d9852-358">String</span><span class="sxs-lookup"><span data-stu-id="d9852-358">String</span></span>|<span data-ttu-id="d9852-359">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="d9852-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-360">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-360">Requirements</span></span>

|<span data-ttu-id="d9852-361">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-361">Requirement</span></span>| <span data-ttu-id="d9852-362">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-364">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-364">1.0</span></span>|
|[<span data-ttu-id="d9852-365">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-366">ReadItem</span></span>|
|[<span data-ttu-id="d9852-367">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d9852-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-368">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-369">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d9852-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d9852-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d9852-371">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="d9852-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-372">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="d9852-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9852-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="d9852-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d9852-p114">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="d9852-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d9852-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="d9852-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d9852-380">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d9852-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-381">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-382">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="d9852-382">All parameters are optional.</span></span>

|<span data-ttu-id="d9852-383">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-383">Name</span></span>| <span data-ttu-id="d9852-384">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-384">Type</span></span>| <span data-ttu-id="d9852-385">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d9852-386">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-386">Object</span></span> | <span data-ttu-id="d9852-387">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="d9852-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d9852-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d9852-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d9852-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d9852-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d9852-394">Date</span><span class="sxs-lookup"><span data-stu-id="d9852-394">Date</span></span> | <span data-ttu-id="d9852-395">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d9852-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d9852-396">Data</span><span class="sxs-lookup"><span data-stu-id="d9852-396">Date</span></span> | <span data-ttu-id="d9852-397">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d9852-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d9852-398">String</span><span class="sxs-lookup"><span data-stu-id="d9852-398">String</span></span> | <span data-ttu-id="d9852-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d9852-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d9852-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d9852-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d9852-404">String</span><span class="sxs-lookup"><span data-stu-id="d9852-404">String</span></span> | <span data-ttu-id="d9852-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d9852-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d9852-407">String</span><span class="sxs-lookup"><span data-stu-id="d9852-407">String</span></span> | <span data-ttu-id="d9852-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d9852-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9852-410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-410">Requirements</span></span>

|<span data-ttu-id="d9852-411">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-411">Requirement</span></span>| <span data-ttu-id="d9852-412">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-414">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-414">1.0</span></span>|
|[<span data-ttu-id="d9852-415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-416">ReadItem</span></span>|
|[<span data-ttu-id="d9852-417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-418">Read</span><span class="sxs-lookup"><span data-stu-id="d9852-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d9852-420">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="d9852-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d9852-421">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d9852-422">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="d9852-423">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="d9852-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d9852-424">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d9852-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-425">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-426">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="d9852-426">All parameters are optional.</span></span>

|<span data-ttu-id="d9852-427">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-427">Name</span></span>| <span data-ttu-id="d9852-428">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-428">Type</span></span>| <span data-ttu-id="d9852-429">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d9852-430">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-430">Object</span></span> | <span data-ttu-id="d9852-431">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d9852-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d9852-433">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="d9852-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="d9852-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d9852-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d9852-436">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="d9852-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="d9852-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d9852-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d9852-439">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="d9852-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="d9852-440">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d9852-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d9852-441">String</span><span class="sxs-lookup"><span data-stu-id="d9852-441">String</span></span> | <span data-ttu-id="d9852-442">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-442">A string containing the subject of the message.</span></span> <span data-ttu-id="d9852-443">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d9852-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d9852-444">String</span><span class="sxs-lookup"><span data-stu-id="d9852-444">String</span></span> | <span data-ttu-id="d9852-445">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-445">The HTML body of the message.</span></span> <span data-ttu-id="d9852-446">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d9852-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d9852-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d9852-448">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="d9852-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d9852-449">String</span><span class="sxs-lookup"><span data-stu-id="d9852-449">String</span></span> | <span data-ttu-id="d9852-p128">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d9852-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d9852-452">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d9852-452">String</span></span> | <span data-ttu-id="d9852-453">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="d9852-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d9852-454">String</span><span class="sxs-lookup"><span data-stu-id="d9852-454">String</span></span> | <span data-ttu-id="d9852-p129">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d9852-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d9852-457">Booliano</span><span class="sxs-lookup"><span data-stu-id="d9852-457">Boolean</span></span> | <span data-ttu-id="d9852-p130">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="d9852-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d9852-460">String</span><span class="sxs-lookup"><span data-stu-id="d9852-460">String</span></span> | <span data-ttu-id="d9852-461">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="d9852-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="d9852-462">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d9852-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="d9852-463">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d9852-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d9852-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-464">Requirements</span></span>

|<span data-ttu-id="d9852-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-465">Requirement</span></span>| <span data-ttu-id="d9852-466">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-468">1.6</span><span class="sxs-lookup"><span data-stu-id="d9852-468">1.6</span></span> |
|[<span data-ttu-id="d9852-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-470">ReadItem</span></span>|
|[<span data-ttu-id="d9852-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-472">Read</span><span class="sxs-lookup"><span data-stu-id="d9852-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d9852-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d9852-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d9852-475">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="d9852-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d9852-p132">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="d9852-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-478">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="d9852-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="d9852-479">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="d9852-479">**REST Tokens**</span></span>

<span data-ttu-id="d9852-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="d9852-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d9852-483">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="d9852-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d9852-484">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="d9852-484">**EWS Tokens**</span></span>

<span data-ttu-id="d9852-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="d9852-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d9852-487">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="d9852-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-488">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-488">Parameters</span></span>

|<span data-ttu-id="d9852-489">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-489">Name</span></span>| <span data-ttu-id="d9852-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-490">Type</span></span>| <span data-ttu-id="d9852-491">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-491">Attributes</span></span>| <span data-ttu-id="d9852-492">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d9852-493">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-493">Object</span></span> | <span data-ttu-id="d9852-494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-494">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-495">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d9852-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d9852-496">Booliano</span><span class="sxs-lookup"><span data-stu-id="d9852-496">Boolean</span></span> |  <span data-ttu-id="d9852-497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-497">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d9852-500">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-500">Object</span></span> |  <span data-ttu-id="d9852-501">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-501">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-502">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d9852-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d9852-503">function</span><span class="sxs-lookup"><span data-stu-id="d9852-503">function</span></span>||<span data-ttu-id="d9852-p136">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-506">Requirements</span></span>

|<span data-ttu-id="d9852-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-507">Requirement</span></span>| <span data-ttu-id="d9852-508">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-510">1,5</span><span class="sxs-lookup"><span data-stu-id="d9852-510">1.5</span></span> |
|[<span data-ttu-id="d9852-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-512">ReadItem</span></span>|
|[<span data-ttu-id="d9852-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-514">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="d9852-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d9852-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d9852-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d9852-517">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d9852-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d9852-p137">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="d9852-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d9852-p138">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d9852-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d9852-523">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d9852-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="d9852-p139">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-526">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-526">Parameters</span></span>

|<span data-ttu-id="d9852-527">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-527">Name</span></span>| <span data-ttu-id="d9852-528">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-528">Type</span></span>| <span data-ttu-id="d9852-529">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-529">Attributes</span></span>| <span data-ttu-id="d9852-530">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d9852-531">function</span><span class="sxs-lookup"><span data-stu-id="d9852-531">function</span></span>||<span data-ttu-id="d9852-p140">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9852-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d9852-534">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-534">Object</span></span>| <span data-ttu-id="d9852-535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-535">&lt;optional&gt;</span></span>|<span data-ttu-id="d9852-536">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d9852-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-537">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-537">Requirements</span></span>

|<span data-ttu-id="d9852-538">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-538">Requirement</span></span>| <span data-ttu-id="d9852-539">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-540">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-541">1.3</span><span class="sxs-lookup"><span data-stu-id="d9852-541">1.3</span></span>|
|[<span data-ttu-id="d9852-542">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-543">ReadItem</span></span>|
|[<span data-ttu-id="d9852-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-545">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="d9852-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-546">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d9852-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d9852-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d9852-548">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="d9852-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d9852-549">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d9852-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-550">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-550">Parameters</span></span>

|<span data-ttu-id="d9852-551">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-551">Name</span></span>| <span data-ttu-id="d9852-552">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-552">Type</span></span>| <span data-ttu-id="d9852-553">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-553">Attributes</span></span>| <span data-ttu-id="d9852-554">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d9852-555">function</span><span class="sxs-lookup"><span data-stu-id="d9852-555">function</span></span>||<span data-ttu-id="d9852-556">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9852-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9852-557">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9852-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d9852-558">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-558">Object</span></span>| <span data-ttu-id="d9852-559">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-559">&lt;optional&gt;</span></span>|<span data-ttu-id="d9852-560">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d9852-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-561">Requirements</span></span>

|<span data-ttu-id="d9852-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-562">Requirement</span></span>| <span data-ttu-id="d9852-563">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-565">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-565">1.0</span></span>|
|[<span data-ttu-id="d9852-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-567">ReadItem</span></span>|
|[<span data-ttu-id="d9852-568">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d9852-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-569">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d9852-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d9852-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d9852-572">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d9852-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-573">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="d9852-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d9852-574">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="d9852-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="d9852-575">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="d9852-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d9852-576">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d9852-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d9852-577">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="d9852-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d9852-578">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="d9852-578">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d9852-579">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="d9852-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d9852-580">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d9852-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d9852-p142">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d9852-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d9852-583">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="d9852-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d9852-584">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="d9852-584">Version differences</span></span>

<span data-ttu-id="d9852-585">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d9852-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d9852-p143">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d9852-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-589">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-589">Parameters</span></span>

|<span data-ttu-id="d9852-590">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-590">Name</span></span>| <span data-ttu-id="d9852-591">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-591">Type</span></span>| <span data-ttu-id="d9852-592">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-592">Attributes</span></span>| <span data-ttu-id="d9852-593">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d9852-594">String</span><span class="sxs-lookup"><span data-stu-id="d9852-594">String</span></span>||<span data-ttu-id="d9852-595">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="d9852-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d9852-596">function</span><span class="sxs-lookup"><span data-stu-id="d9852-596">function</span></span>||<span data-ttu-id="d9852-597">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9852-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9852-598">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9852-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d9852-599">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="d9852-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d9852-600">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-600">Object</span></span>| <span data-ttu-id="d9852-601">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-601">&lt;optional&gt;</span></span>|<span data-ttu-id="d9852-602">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d9852-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-603">Requirements</span></span>

|<span data-ttu-id="d9852-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-604">Requirement</span></span>| <span data-ttu-id="d9852-605">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-607">1.0</span><span class="sxs-lookup"><span data-stu-id="d9852-607">1.0</span></span>|
|[<span data-ttu-id="d9852-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d9852-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d9852-610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-611">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9852-612">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d9852-612">Example</span></span>

<span data-ttu-id="d9852-613">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="d9852-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d9852-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9852-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d9852-615">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="d9852-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d9852-616">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="d9852-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9852-617">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d9852-617">Parameters</span></span>

| <span data-ttu-id="d9852-618">Nome</span><span class="sxs-lookup"><span data-stu-id="d9852-618">Name</span></span> | <span data-ttu-id="d9852-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="d9852-619">Type</span></span> | <span data-ttu-id="d9852-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="d9852-620">Attributes</span></span> | <span data-ttu-id="d9852-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="d9852-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d9852-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d9852-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d9852-623">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="d9852-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d9852-624">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-624">Object</span></span> | <span data-ttu-id="d9852-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-625">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-626">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d9852-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d9852-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="d9852-627">Object</span></span> | <span data-ttu-id="d9852-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-628">&lt;optional&gt;</span></span> | <span data-ttu-id="d9852-629">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d9852-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d9852-630">function</span><span class="sxs-lookup"><span data-stu-id="d9852-630">function</span></span>| <span data-ttu-id="d9852-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9852-631">&lt;optional&gt;</span></span>|<span data-ttu-id="d9852-632">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9852-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9852-633">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d9852-633">Requirements</span></span>

|<span data-ttu-id="d9852-634">Requisito</span><span class="sxs-lookup"><span data-stu-id="d9852-634">Requirement</span></span>| <span data-ttu-id="d9852-635">Valor</span><span class="sxs-lookup"><span data-stu-id="d9852-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9852-636">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d9852-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9852-637">1,5</span><span class="sxs-lookup"><span data-stu-id="d9852-637">1.5</span></span> |
|[<span data-ttu-id="d9852-638">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d9852-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9852-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9852-639">ReadItem</span></span> |
|[<span data-ttu-id="d9852-640">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d9852-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d9852-641">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d9852-641">Compose or Read</span></span>|
